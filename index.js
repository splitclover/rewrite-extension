import { sendOpenAIRequest, oai_settings } from "../../../openai.js";
import { eventSource, event_types, saveSettingsDebounced, messageFormatting, addCopyToCodeBlocks } from "../../../../script.js";
import { extension_settings, getContext } from "../../../extensions.js";

const extensionName = "rewrite-extension";
const extensionFolderPath = `scripts/extensions/third-party/${extensionName}`;

// Default settings
const defaultSettings = {
    rewritePreset: "",
    shortenPreset: "",
    expandPreset: "",
    highlightDuration: 3000
};

let rewriteMenu = null;
let lastSelection = null;
let abortController;

let lastChange = {
    mesId: null,
    swipeId: null,
    originalContent: null,
    messageDiv: null
};

// Load settings
function loadSettings() {
    extension_settings[extensionName] = extension_settings[extensionName] || {};
    if (Object.keys(extension_settings[extensionName]).length === 0) {
        Object.assign(extension_settings[extensionName], defaultSettings);
    }

    // Ensure highlightDuration has a value
    if (extension_settings[extensionName].highlightDuration === undefined) {
        extension_settings[extensionName].highlightDuration = defaultSettings.highlightDuration;
    }

    $("#rewrite_preset").val(extension_settings[extensionName].rewritePreset);
    $("#shorten_preset").val(extension_settings[extensionName].shortenPreset);
    $("#expand_preset").val(extension_settings[extensionName].expandPreset);
    $("#highlight_duration").val(extension_settings[extensionName].highlightDuration);
}

// Save settings
function saveSettings() {
    extension_settings[extensionName].rewritePreset = $("#rewrite_preset").val();
    extension_settings[extensionName].shortenPreset = $("#shorten_preset").val();
    extension_settings[extensionName].expandPreset = $("#expand_preset").val();
    extension_settings[extensionName].highlightDuration = parseInt($("#highlight_duration").val()) || defaultSettings.highlightDuration;
    saveSettingsDebounced();
}

// Populate dropdowns
async function populateDropdowns() {
    const result = await fetch('/api/settings/get', {
        method: 'POST',
        headers: getContext().getRequestHeaders(),
        body: JSON.stringify({}),
    });

    if (result.ok) {
        const data = await result.json();
        const presets = data.openai_setting_names;

        const dropdowns = ['rewrite_preset', 'shorten_preset', 'expand_preset'];
        dropdowns.forEach(dropdown => {
            const select = $(`#${dropdown}`);
            select.empty();
            presets.forEach(preset => {
                select.append($('<option>', {
                    value: preset,
                    text: preset
                }));
            });
        });

        // Set the selected values after populating
        loadSettings();
    }
}

// Initialize
jQuery(async () => {
    const settingsHtml = await $.get(`${extensionFolderPath}/rewrite_settings.html`);
    $("#extensions_settings2").append(settingsHtml);

    // Populate dropdowns
    await populateDropdowns();

    // Add event listeners
    $(".rewrite-extension-settings select, #highlight_duration").on("change", saveSettings);

    // Load settings
    loadSettings();

    // Add event listener for SETTINGS_UPDATED
    eventSource.on(event_types.SETTINGS_UPDATED, () => {
        populateDropdowns();
    });

    eventSource.on(event_types.MESSAGE_EDITED, (editedMesId) => {
        if (lastChange.mesId === editedMesId) {
            removeUndoButton(editedMesId);
            lastChange = { mesId: null, swipeId: null, originalContent: null, messageDiv: null };
        }
    });
});

// Initialize the rewrite menu functionality
initRewriteMenu();


function initRewriteMenu() {
    // document.addEventListener('mouseup', handleSelectionEnd);
    // document.addEventListener('touchend', handleSelectionEnd);
    document.addEventListener('selectionchange', handleSelectionChange);
    document.addEventListener('mousedown', hideMenuOnOutsideClick);
    document.addEventListener('touchstart', hideMenuOnOutsideClick);

    let chatContainer = document.getElementById('chat');
    chatContainer.addEventListener('scroll', positionMenu);

    $('#mes_stop').on('click', handleStopRewrite);
}


function handleStopRewrite() {
    if (abortController) {
        const { prev_oai_settings, mesDiv, mesId, swipeId, highlightDuration } = abortController.signal;
        abortController.abort();
        // Restore the original settings
        Object.assign(oai_settings, prev_oai_settings);
        getContext().activateSendButtons();

        // Call removeHighlight with the stored arguments
        setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);
    }
}

// function handleSelectionEnd(e) {
//     if (e.target && e.target.closest('.ctx-menu')) return;
//     removeRewriteMenu();
//     setTimeout(processSelection, 50);
// }

function handleSelectionChange() {
    // Use a small timeout to ensure the selection has been updated
    setTimeout(processSelection, 50);
}

function processSelection() {
    // First, check if getContext().chatId is defined
    if (getContext().chatId === undefined) {
        return; // Exit the function if chatId is undefined
    }

    let selection = window.getSelection();
    let selectedText = selection.toString().trim();

    // Always remove the existing menu first
    removeRewriteMenu();

    if (selectedText.length > 0) {
        let range = selection.getRangeAt(0);

        // Find the mes_text elements for both start and end of the selection
        let startMesText = range.startContainer.nodeType === Node.ELEMENT_NODE
            ? range.startContainer.closest('.mes_text')
            : range.startContainer.parentElement.closest('.mes_text');

        let endMesText = range.endContainer.nodeType === Node.ELEMENT_NODE
            ? range.endContainer.closest('.mes_text')
            : range.endContainer.parentElement.closest('.mes_text');

        // Check if both start and end are within the same mes_text element
        if (startMesText && endMesText && startMesText === endMesText) {
            createRewriteMenu();
        }
    }

    lastSelection = selectedText.length > 0 ? selectedText : null;
}

async function handleMenuItemClick(e) {
    e.preventDefault();
    e.stopPropagation();

    const option = e.target.dataset.option;
    const selection = window.getSelection();
    const selectedText = selection.toString().trim();

    if (selectedText) {
        const mesTextElement = findClosestMesText(selection.anchorNode);
        if (mesTextElement) {
            const messageDiv = findMessageDiv(mesTextElement);
            if (messageDiv) {
                const mesId = messageDiv.getAttribute('mesid');
                const swipeId = messageDiv.getAttribute('swipeid');

                // console.log(`${option} option clicked!`);
                // console.log('Message ID:', mesId);
                // console.log('Swipe ID:', swipeId);
                // toastr.info(`${option} option clicked! Message ID: ${mesId}, Swipe ID: ${swipeId}`);

                await handleRewrite(mesId, swipeId, option);
            }
        }
    }

    removeRewriteMenu();

    window.getSelection().removeAllRanges();
}

function hideMenuOnOutsideClick(e) {
    if (rewriteMenu && !rewriteMenu.contains(e.target)) {
        removeRewriteMenu();
    }
}

function createRewriteMenu() {
    removeRewriteMenu();

    rewriteMenu = document.createElement('ul');
    rewriteMenu.className = 'list-group ctx-menu';
    rewriteMenu.style.position = 'absolute';
    rewriteMenu.style.zIndex = '1000';
    rewriteMenu.style.position = 'fixed';

    const options = ['Rewrite', 'Shorten', 'Expand'];
    options.forEach(option => {
        let li = document.createElement('li');
        li.className = 'list-group-item ctx-item';
        li.textContent = option;
        li.addEventListener('mousedown', handleMenuItemClick);
        li.addEventListener('touchstart', handleMenuItemClick);
        li.dataset.option = option;
        rewriteMenu.appendChild(li);
    });

    document.body.appendChild(rewriteMenu);
    positionMenu();
}

function positionMenu() {
    if (!rewriteMenu) return;

    let selection = window.getSelection();
    let range = selection.getRangeAt(0);
    let rect = range.getBoundingClientRect();

    // Calculate the menu's position
    let left = rect.left + window.pageXOffset;
    let top = rect.bottom + window.pageYOffset + 5;

    // Get the viewport dimensions
    let viewportWidth = window.innerWidth;
    let viewportHeight = window.innerHeight;

    // Get the menu's dimensions
    let menuWidth = rewriteMenu.offsetWidth;
    let menuHeight = rewriteMenu.offsetHeight;

    // Adjust the position if the menu overflows the viewport
    if (left + menuWidth > viewportWidth) {
        left = viewportWidth - menuWidth;
    }
    if (top + menuHeight > viewportHeight) {
        top = rect.top + window.pageYOffset - menuHeight - 5;
    }

    rewriteMenu.style.left = `${left}px`;
    rewriteMenu.style.top = `${top}px`;
}

function removeRewriteMenu() {
    if (rewriteMenu) {
        rewriteMenu.remove();
        rewriteMenu = null;
    }
}

function addUndoButton() {
    if (lastChange.messageDiv) {
        const mesButtons = lastChange.messageDiv.querySelector('.mes_buttons');
        if (mesButtons) {
            const undoButton = document.createElement('div');
            undoButton.className = 'mes_button mes_undo_rewrite fa-solid fa-undo interactable';
            undoButton.title = 'Undo rewrite';
            undoButton.tabIndex = 0;
            undoButton.addEventListener('click', handleUndo);

            // Insert as the second child
            if (mesButtons.children.length >= 1) {
                mesButtons.insertBefore(undoButton, mesButtons.children[1]);
            } else {
                mesButtons.appendChild(undoButton);
            }
        }
    }
}

function removeUndoButton(mesId) {
    const messageDiv = document.querySelector(`[mesid="${mesId}"]`);
    if (messageDiv) {
        const undoButton = messageDiv.querySelector('.mes_undo_rewrite');
        if (undoButton) {
            undoButton.remove();
        }
    }
}

async function removeHighlight(mesDiv, mesId, swipeId) {
    const highlightSpan = mesDiv.querySelector('.animated-highlight');
    if (highlightSpan) {
        const textNode = document.createTextNode(highlightSpan.textContent);
        highlightSpan.parentNode.replaceChild(textNode, highlightSpan);
    }

    const context = getContext();
    const messageData = context.chat[mesId];

    if (messageData) {
        let messageContent;
        if (swipeId !== undefined && messageData.swipes && messageData.swipes[swipeId]) {
            messageContent = messageData.swipes[swipeId];
        } else {
            messageContent = messageData.mes;
        }

        // Format the message into HTML
        const formattedMessage = messageFormatting(
            messageContent,
            context.name2,
            messageData.isSystem,
            messageData.isUser,
            mesId
        );

        // Create a temporary div to hold the formatted message
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = formattedMessage;

        // Apply addCopyToCodeBlocks to the temporary div
        addCopyToCodeBlocks(tempDiv);

        // Find the mes_text element within the message div
        const mesTextElement = mesDiv.closest('.mes').querySelector('.mes_text');
        if (mesTextElement) {
            // Replace the content of mes_text with the new formatted content
            mesTextElement.innerHTML = tempDiv.innerHTML;
        }
    }
}

function findClosestMesText(element) {
    while (element && element.nodeType !== 1) {
        element = element.parentElement;
    }
    while (element) {
        if (element.classList && element.classList.contains('mes_text')) {
            return element;
        }
        element = element.parentElement;
    }
    return null;
}

function findMessageDiv(element) {
    while (element) {
        if (element.hasAttribute('mesid') && element.hasAttribute('swipeid')) {
            return element;
        }
        element = element.parentElement;
    }
    return null;
}

function createTextMapping(rawText, formattedHtml) {
    const formattedText = stripHtml(formattedHtml);
    const mapping = [];
    let rawIndex = 0;
    let formattedIndex = 0;

    while (rawIndex < rawText.length && formattedIndex < formattedText.length) {
        if (rawText[rawIndex] === formattedText[formattedIndex]) {
            mapping.push([rawIndex, formattedIndex]);
            rawIndex++;
            formattedIndex++;
        } else if (rawText.substr(rawIndex, 3) === '...' && formattedText[formattedIndex] === '…') {
            // Handle ellipsis
            mapping.push([rawIndex, formattedIndex]);
            mapping.push([rawIndex + 1, formattedIndex]);
            mapping.push([rawIndex + 2, formattedIndex]);
            rawIndex += 3;
            formattedIndex++;
        } else if (formattedText[formattedIndex] === ' ' || formattedText[formattedIndex] === '\n') {
            // Skip extra whitespace in formatted text
            formattedIndex++;
        } else {
            // Skip characters in raw text that don't appear in formatted text
            rawIndex++;
        }
    }

    return {
        formattedToRaw: (formattedOffset) => {
            let low = 0;
            let high = mapping.length - 1;

            while (low <= high) {
                let mid = Math.floor((low + high) / 2);
                if (mapping[mid][1] === formattedOffset) {
                    return mapping[mid][0];
                } else if (mapping[mid][1] < formattedOffset) {
                    low = mid + 1;
                } else {
                    high = mid - 1;
                }
            }

            // If we didn't find an exact match, return the closest one
            if (low > 0) low--;
            return mapping[low][0] + (formattedOffset - mapping[low][1]);
        }
    };
}

function stripHtml(html) {
    const tmp = document.createElement('DIV');
    tmp.innerHTML = html;
    return tmp.textContent || tmp.innerText || '';
}

function getTextOffset(parent, node) {
    const treeWalker = document.createTreeWalker(
        parent,
        NodeFilter.SHOW_TEXT,
        null,
        false
    );

    let offset = 0;
    while (treeWalker.nextNode() !== node) {
        offset += treeWalker.currentNode.length;
    }

    return offset;
}

async function handleRewrite(mesId, swipeId, option) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    if (!mesDiv) return; // Exit if we can't find the message div

    const selection = window.getSelection();
    const range = selection.getRangeAt(0);

    // Get the full message content
    const fullMessage = getContext().chat[mesId].mes;

    // Remove the undo button for the previous rewrite, if it exists
    if (lastChange.mesId) {
        removeUndoButton(lastChange.mesId);
    }
    // For undo
    lastChange.mesId = mesId;
    lastChange.swipeId = swipeId;
    lastChange.originalContent = fullMessage;
    lastChange.messageDiv = document.querySelector(`[mesid="${mesId}"]`);

    // Get the formatted message
    const formattedMessage = messageFormatting(fullMessage, undefined, getContext().chat[mesId].isSystem, getContext().chat[mesId].isUser, mesId);

    // Create a mapping between raw and formatted text
    const mapping = createTextMapping(fullMessage, formattedMessage);

    // Calculate the start and end offsets relative to the formatted text content
    const startOffset = getTextOffset(mesDiv, range.startContainer) + range.startOffset;
    const endOffset = getTextOffset(mesDiv, range.endContainer) + range.endOffset;

    // Map these offsets back to the raw message
    const rawStartOffset = mapping.formattedToRaw(startOffset);
    const rawEndOffset = mapping.formattedToRaw(endOffset);

    // Get the selected raw text
    const selectedRawText = fullMessage.substring(rawStartOffset, rawEndOffset);
    // console.log(rawStartOffset);
    // console.log(rawEndOffset);


    // Get the selected preset based on the option
    let selectedPreset;
    switch (option) {
        case 'Rewrite':
            selectedPreset = extension_settings[extensionName].rewritePreset;
            break;
        case 'Shorten':
            selectedPreset = extension_settings[extensionName].shortenPreset;
            break;
        case 'Expand':
            selectedPreset = extension_settings[extensionName].expandPreset;
            break;
        default:
            return; // Exit if the option is not recognized
    }

    // Fetch the settings
    const result = await fetch('/api/settings/get', {
        method: 'POST',
        headers: getContext().getRequestHeaders(),
        body: JSON.stringify({}),
    });

    if (!result.ok) {
        console.error('Failed to fetch settings');
        return;
    }

    const data = await result.json();
    const presetIndex = data.openai_setting_names.indexOf(selectedPreset);
    if (presetIndex === -1) {
        console.error('Selected preset not found');
        return;
    }

    // Save the current settings
    const prev_oai_settings = Object.assign({}, oai_settings);

    // Parse the selected preset settings
    let selectedPresetSettings;
    try {
        selectedPresetSettings = JSON.parse(data.openai_settings[presetIndex]);
    } catch (error) {
        console.error('Error parsing preset settings:', error);
        return;
    }

    // Override oai_settings with the selected preset
    Object.assign(oai_settings, selectedPresetSettings);

    // Set up the event listener for the generated prompt
    const promptReadyPromise = new Promise(resolve => {
        eventSource.once(event_types.CHAT_COMPLETION_PROMPT_READY, resolve);
    });

    // Generate the prompt
    getContext().generate('normal', {}, true);

    // Wait for the prompt to be ready
    const promptData = await promptReadyPromise;

    // Substitute {{rewrite}} macro with the selected text directly in the promptData.chat
    promptData.chat = promptData.chat.map(message => {
        if (Array.isArray(message.content)) {
            // If content is an array, process only the text entries
            message.content = message.content.map(item => {
                if (item.type === 'text') {
                    item.text = item.text.replace(/{{rewrite}}/g, selectedRawText);
                    item.text = item.text.replace(/{{targetmessage}}/g, fullMessage);
                }
                return item;
            });
        } else if (typeof message.content === 'string') {
            // If content is a string, process it directly
            message.content = message.content.replace(/{{rewrite}}/g, selectedRawText);
            message.content = message.content.replace(/{{targetmessage}}/g, fullMessage);
        }
        return message;
    });

    // Create a new AbortController
    abortController = new AbortController();

    // Store the necessary data in the signal
    abortController.signal.prev_oai_settings = prev_oai_settings;
    abortController.signal.mesDiv = mesDiv;
    abortController.signal.mesId = mesId;
    abortController.signal.swipeId = swipeId;
    abortController.signal.highlightDuration = extension_settings[extensionName].highlightDuration;

    // Show the stop button
    getContext().deactivateSendButtons();

    const res = await sendOpenAIRequest('normal', promptData.chat, abortController.signal);
    window.getSelection().removeAllRanges();

    let newText = '';

    if (typeof res === 'function') {
        // Streaming case

        const streamingSpan = document.createElement('span');
        streamingSpan.className = 'animated-highlight';

        // Replace the selected text with the streaming span
        range.deleteContents();
        range.insertNode(streamingSpan);

        for await (const chunk of res()) {
            newText = chunk.text;
            streamingSpan.textContent = newText;
        }
    } else {
        // Non-streaming case
        newText = res?.choices?.[0]?.message?.content ?? res?.choices?.[0]?.text ?? res?.text ?? '';
        const highlightedNewText = document.createElement('span');
        highlightedNewText.className = 'animated-highlight';
        highlightedNewText.textContent = newText;

        range.deleteContents();
        range.insertNode(highlightedNewText);
    }

    // Restore the original settings
    Object.assign(oai_settings, prev_oai_settings);

    // Remove highlight after x seconds when streaming is complete
    const highlightDuration = extension_settings[extensionName].highlightDuration;
    setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

    // Undo button
    addUndoButton();

    getContext().activateSendButtons();
    await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);
}

async function handleUndo() {
    if (lastChange.mesId && lastChange.originalContent) {
        const context = getContext();
        const messageDiv = document.querySelector(`[mesid="${lastChange.mesId}"]`);

        if (!messageDiv || !context.chat[lastChange.mesId]) {
            console.error('Message not found for undo operation');
            return;
        }

        // Update the chat context
        context.chat[lastChange.mesId].mes = lastChange.originalContent;

        // Only update swipes if they exist
        if (lastChange.swipeId !== undefined && context.chat[lastChange.mesId].swipes) {
            context.chat[lastChange.mesId].swipes[lastChange.swipeId] = lastChange.originalContent;
        }

        // Update the UI
        const mesTextElement = messageDiv.querySelector('.mes_text');
        if (mesTextElement) {
            mesTextElement.innerHTML = messageFormatting(
                lastChange.originalContent,
                context.name2,
                context.chat[lastChange.mesId].isSystem,
                context.chat[lastChange.mesId].isUser,
                lastChange.mesId
            );
            addCopyToCodeBlocks(mesTextElement);
        }

        // Save the chat
        await context.saveChat();

        // Remove the undo button
        const undoButton = messageDiv.querySelector('.mes_undo_rewrite');
        if (undoButton) {
            undoButton.remove();
        }

        // Clear the last change
        lastChange = { mesId: null, swipeId: null, originalContent: null };
    }
}

async function saveRewrittenText(mesId, swipeId, fullMessage, startOffset, endOffset, newText) {
    const context = getContext();

    // Create the new message with the rewritten section
    const newMessage =
        fullMessage.substring(0, startOffset) +
        newText +
        fullMessage.substring(endOffset);

    // Update the main message
    context.chat[mesId].mes = newMessage;

    // Update the swipe if it exists
    if (swipeId !== undefined && context.chat[mesId].swipes && context.chat[mesId].swipes[swipeId]) {
        context.chat[mesId].swipes[swipeId] = newMessage;
    }

    // Save and reload the chat
    await context.saveChat();
}
