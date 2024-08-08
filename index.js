import { sendOpenAIRequest, oai_settings } from "../../../openai.js";
import { extractAllWords } from "../../../utils.js";
import { getTokenCount } from "../../../tokenizers.js";
import { getNovelGenerationData, generateNovelWithStreaming, nai_settings } from "../../../nai-settings.js";
import { getTextGenGenerationData, generateTextGenWithStreaming } from "../../../textgen-settings.js";
import {
    main_api,
    novelai_settings,
    novelai_setting_names,
    eventSource,
    event_types,
    saveSettingsDebounced,
    messageFormatting,
    addCopyToCodeBlocks,
    getRequestHeaders,
} from "../../../../script.js";
import { extension_settings, getContext } from "../../../extensions.js";

const extensionName = "rewrite-extension";
const extensionFolderPath = `scripts/extensions/third-party/${extensionName}`;

const undo_steps = 15;

// Default settings
const defaultSettings = {
    rewritePreset: "",
    shortenPreset: "",
    expandPreset: "",
    highlightDuration: 3000,
    selectedModel: "chat_completion",
    textRewritePrompt: `[INST]Rewrite this section of text: """{{rewrite}}""" while keeping the same content, general style and length. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the rewritten text without any comments: `,
    textShortenPrompt: `[INST]Rewrite this section of text: """{{rewrite}}""" while keeping the same content, general style. Do not list alternatives and only print the result without prefix or suffix. Shorten it by roughly 20%.[/INST]

Sure, here is only the rewritten text without any comments: `,
    textExpandPrompt: `[INST]Rewrite this section of text: """{{rewrite}}""" while keeping the same content, general style. Do not list alternatives and only print the result without prefix or suffix. Lengthen it by roughly 20%.[/INST]

Sure, here is only the rewritten text without any comments: `,
    useStreaming: true,
    useDynamicTokens: true,
    dynamicTokenMode: 'multiplicative',
    rewriteTokens: 100,
    shortenTokens: 50,
    expandTokens: 150,
    rewriteTokensAdd: 0,
    shortenTokensAdd: -50,
    expandTokensAdd: 50,
    rewriteTokensMult: 1.05,
    shortenTokensMult: 0.8,
    expandTokensMult: 1.5,
    removePrefix: `"`,
    removeSuffix: `"`,
    overrideMaxTokens: true,
    showRewrite: true,
    showShorten: true,
    showExpand: true,
    showDelete: true,
};

let rewriteMenu = null;
let lastSelection = null;
let abortController;

let changeHistory = [];

// Load settings
function loadSettings() {
    extension_settings[extensionName] = extension_settings[extensionName] || {};

    // Helper function to get a setting with a default value
    const getSetting = (key, defaultValue) => {
        return extension_settings[extensionName][key] !== undefined
            ? extension_settings[extensionName][key]
            : defaultValue;
    };

    // Load settings, using defaults if not set
    $("#rewrite_preset").val(getSetting('rewritePreset', defaultSettings.rewritePreset));
    $("#shorten_preset").val(getSetting('shortenPreset', defaultSettings.shortenPreset));
    $("#expand_preset").val(getSetting('expandPreset', defaultSettings.expandPreset));
    $("#highlight_duration").val(getSetting('highlightDuration', defaultSettings.highlightDuration));
    $("#rewrite_extension_model_select").val(getSetting('selectedModel', defaultSettings.selectedModel));
    $("#text_rewrite_prompt").val(getSetting('textRewritePrompt', defaultSettings.textRewritePrompt));
    $("#text_shorten_prompt").val(getSetting('textShortenPrompt', defaultSettings.textShortenPrompt));
    $("#text_expand_prompt").val(getSetting('textExpandPrompt', defaultSettings.textExpandPrompt));
    $("#use_streaming").prop('checked', getSetting('useStreaming', defaultSettings.useStreaming));
    $("#use_dynamic_tokens").prop('checked', getSetting('useDynamicTokens', defaultSettings.useDynamicTokens));
    $("#dynamic_token_mode").val(getSetting('dynamicTokenMode', defaultSettings.dynamicTokenMode));
    $("#rewrite_tokens").val(getSetting('rewriteTokens', defaultSettings.rewriteTokens));
    $("#shorten_tokens").val(getSetting('shortenTokens', defaultSettings.shortenTokens));
    $("#expand_tokens").val(getSetting('expandTokens', defaultSettings.expandTokens));
    $("#rewrite_tokens_add").val(getSetting('rewriteTokensAdd', defaultSettings.rewriteTokensAdd));
    $("#shorten_tokens_add").val(getSetting('shortenTokensAdd', defaultSettings.shortenTokensAdd));
    $("#expand_tokens_add").val(getSetting('expandTokensAdd', defaultSettings.expandTokensAdd));
    $("#rewrite_tokens_mult").val(getSetting('rewriteTokensMult', defaultSettings.rewriteTokensMult));
    $("#shorten_tokens_mult").val(getSetting('shortenTokensMult', defaultSettings.shortenTokensMult));
    $("#expand_tokens_mult").val(getSetting('expandTokensMult', defaultSettings.expandTokensMult));
    $("#remove_prefix").val(getSetting('removePrefix', defaultSettings.removePrefix));
    $("#remove_suffix").val(getSetting('removeSuffix', defaultSettings.removeSuffix));
    $("#override_max_tokens").prop('checked', getSetting('overrideMaxTokens', defaultSettings.overrideMaxTokens));
    $("#show_rewrite").prop('checked', getSetting('showRewrite', defaultSettings.showRewrite));
    $("#show_shorten").prop('checked', getSetting('showShorten', defaultSettings.showShorten));
    $("#show_expand").prop('checked', getSetting('showExpand', defaultSettings.showExpand));
    $("#show_delete").prop('checked', getSetting('showDelete', defaultSettings.showDelete));

    // Update the UI based on loaded settings
    updateModelSettings();
    updateTokenSettings();
}

function saveSettings() {
    extension_settings[extensionName] = {
        rewritePreset: $("#rewrite_preset").val(),
        shortenPreset: $("#shorten_preset").val(),
        expandPreset: $("#expand_preset").val(),
        highlightDuration: parseInt($("#highlight_duration").val()),
        selectedModel: $("#rewrite_extension_model_select").val(),
        textRewritePrompt: $("#text_rewrite_prompt").val(),
        textShortenPrompt: $("#text_shorten_prompt").val(),
        textExpandPrompt: $("#text_expand_prompt").val(),
        useStreaming: $("#use_streaming").is(':checked'),
        useDynamicTokens: $("#use_dynamic_tokens").is(':checked'),
        dynamicTokenMode: $("#dynamic_token_mode").val(),
        rewriteTokens: parseInt($("#rewrite_tokens").val()),
        shortenTokens: parseInt($("#shorten_tokens").val()),
        expandTokens: parseInt($("#expand_tokens").val()),
        rewriteTokensAdd: parseInt($("#rewrite_tokens_add").val()),
        shortenTokensAdd: parseInt($("#shorten_tokens_add").val()),
        expandTokensAdd: parseInt($("#expand_tokens_add").val()),
        rewriteTokensMult: parseFloat($("#rewrite_tokens_mult").val()),
        shortenTokensMult: parseFloat($("#shorten_tokens_mult").val()),
        expandTokensMult: parseFloat($("#expand_tokens_mult").val()),
        removePrefix: $("#remove_prefix").val(),
        removeSuffix: $("#remove_suffix").val(),
        overrideMaxTokens: $("#override_max_tokens").is(':checked'),
        showRewrite: $("#show_rewrite").is(':checked'),
        showShorten: $("#show_shorten").is(':checked'),
        showExpand: $("#show_expand").is(':checked'),
        showDelete: $("#show_delete").is(':checked'),
    };

    // Ensure all settings have a value, using defaults if necessary
    for (const [key, value] of Object.entries(defaultSettings)) {
        if (extension_settings[extensionName][key] === undefined) {
            extension_settings[extensionName][key] = value;
        }
    }

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

function updateModelSettings() {
    const modelSelect = document.getElementById('rewrite_extension_model_select');
    const chatCompletionSettings = document.getElementById('chat_completion_settings');
    const textBasedSettings = document.getElementById('text_based_settings');

    if (modelSelect.value === 'chat_completion') {
        chatCompletionSettings.style.display = 'block';
        textBasedSettings.style.display = 'none';
    } else {
        chatCompletionSettings.style.display = 'none';
        textBasedSettings.style.display = 'block';
    }
}

function updateTokenSettings() {
    const useDynamicTokens = $("#use_dynamic_tokens").is(':checked');
    const dynamicTokenMode = $("#dynamic_token_mode").val();
    $("#static_token_settings").toggle(!useDynamicTokens);
    $("#dynamic_token_settings").toggle(useDynamicTokens);
    $("#additive_settings").toggle(dynamicTokenMode === 'additive');
    $("#multiplicative_settings").toggle(dynamicTokenMode === 'multiplicative');
}

// Initialize
jQuery(async () => {
    const settingsHtml = await $.get(`${extensionFolderPath}/rewrite_settings.html`);
    $("#extensions_settings2").append(settingsHtml);

    // Populate dropdowns
    await populateDropdowns();

    // Add event listeners
    $(".rewrite-extension-settings select, #highlight_duration, #text_rewrite_prompt, #text_shorten_prompt, #text_expand_prompt").on("change", saveSettings);
    $("#use_streaming").on("change", saveSettings);
    $("#use_dynamic_tokens, #dynamic_token_mode").on("change", () => {
        updateTokenSettings();
        saveSettings();
    });
    $("#rewrite_tokens, #shorten_tokens, #expand_tokens, #rewrite_tokens_add, #shorten_tokens_add, #expand_tokens_add, #rewrite_tokens_mult, #shorten_tokens_mult, #expand_tokens_mult").on("input", saveSettings);
    $("#remove_prefix, #remove_suffix").on("change", saveSettings);
    $("#override_max_tokens").on("change", saveSettings);
    $("#show_rewrite, #show_shorten, #show_expand, #show_delete").on("change", saveSettings);

    $("#rewrite_extension_model_select").on("change", () => {
        updateModelSettings();
        saveSettings();
    });

    // Load settings
    loadSettings();

    // Add event listener for SETTINGS_UPDATED
    eventSource.on(event_types.SETTINGS_UPDATED, () => {
        populateDropdowns();
    });

    eventSource.on(event_types.CHAT_CHANGED, () => {
        changeHistory = [];
        updateUndoButtons();
    });

    eventSource.on(event_types.MESSAGE_EDITED, (editedMesId) => {
        removeUndoButton(editedMesId);
    });

    updateModelSettings();
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
        const { mesDiv, mesId, swipeId, highlightDuration } = abortController.signal;
        abortController.abort();
        // Restore the original settings
        if (abortController.signal.prev_oai_settings) {
            Object.assign(oai_settings, abortController.signal.prev_oai_settings);
        }

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

                if (option === 'Delete') {
                    await handleDeleteSelection(mesId, swipeId);
                } else {
                    await handleRewrite(mesId, swipeId, option);
                }
            }
        }
    }

    removeRewriteMenu();
    window.getSelection().removeAllRanges();
}

async function handleDeleteSelection(mesId, swipeId) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = getSelectedTextInfo(mesId, mesDiv);

    // Create the new message with the deleted section removed
    const newMessage = fullMessage.slice(0, rawStartOffset) + fullMessage.slice(rawEndOffset);

    // Save the change to the history (this also calls updateUndoButtons)
    saveLastChange(mesId, swipeId, fullMessage, newMessage);

    // Update the message in the chat context
    getContext().chat[mesId].mes = newMessage;
    if (swipeId !== undefined && getContext().chat[mesId].swipes) {
        getContext().chat[mesId].swipes[swipeId] = newMessage;
    }

    // Update the UI
    mesDiv.innerHTML = messageFormatting(newMessage, getContext().name2, getContext().chat[mesId].isSystem, getContext().chat[mesId].isUser, mesId);
    addCopyToCodeBlocks(mesDiv);

    // Save the chat
    await getContext().saveChat();
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

    const options = [
        { name: 'Rewrite', show: extension_settings[extensionName].showRewrite },
        { name: 'Shorten', show: extension_settings[extensionName].showShorten },
        { name: 'Expand', show: extension_settings[extensionName].showExpand },
        { name: 'Delete', show: extension_settings[extensionName].showDelete }
    ];
    options.forEach(option => {
        if (option.show) {
            let li = document.createElement('li');
            li.className = 'list-group-item ctx-item';
            li.textContent = option.name;
            li.addEventListener('mousedown', handleMenuItemClick);
            li.addEventListener('touchstart', handleMenuItemClick);
            li.dataset.option = option.name;
            rewriteMenu.appendChild(li);
        }
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

function addUndoButton(mesId) {
    const messageDiv = document.querySelector(`[mesid="${mesId}"]`);
    if (messageDiv) {
        const mesButtons = messageDiv.querySelector('.mes_buttons');
        if (mesButtons) {
            const undoButton = document.createElement('div');
            undoButton.className = 'mes_button mes_undo_rewrite fa-solid fa-undo interactable';
            undoButton.title = 'Undo rewrite';
            undoButton.dataset.mesId = mesId;
            undoButton.addEventListener('click', handleUndo);

            if (mesButtons.children.length >= 1) {
                mesButtons.insertBefore(undoButton, mesButtons.children[1]);
            } else {
                mesButtons.appendChild(undoButton);
            }
        }
    }
}

function removeUndoButton(editedMesId) {
    // Remove all changes for this message from the changeHistory
    changeHistory = changeHistory.filter(change => change.mesId !== editedMesId);

    // Update undo buttons for other messages
    updateUndoButtons();
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

function getSelectedTextInfo(mesId, mesDiv) {
    const selection = window.getSelection();
    const range = selection.getRangeAt(0);

    // Get the full message content
    const fullMessage = getContext().chat[mesId].mes;

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

    return {
        fullMessage,
        selectedRawText,
        rawStartOffset,
        rawEndOffset,
        range
    };
}

function saveLastChange(mesId, swipeId, originalContent, newContent) {
    changeHistory.push({
        mesId,
        swipeId,
        originalContent,
        newContent,
        timestamp: Date.now()
    });

    // Limit history to last n changes
    if (changeHistory.length > undo_steps) {
        changeHistory.shift();
    }

    updateUndoButtons();
}

function updateUndoButtons() {
    // Remove all existing undo buttons
    document.querySelectorAll('.mes_undo_rewrite').forEach(button => button.remove());

    // Add undo buttons for all messages with changes
    const changedMessageIds = [...new Set(changeHistory.map(change => change.mesId))];
    changedMessageIds.forEach(mesId => addUndoButton(mesId));
}

async function handleRewrite(mesId, swipeId, option) {
    if (main_api === 'openai') {
        const selectedModel = extension_settings[extensionName].selectedModel;
        if (selectedModel === 'chat_completion') {
            return handleChatCompletionRewrite(mesId, swipeId, option);
        } else {
            return handleSimplifiedChatCompletionRewrite(mesId, swipeId, option);
        }
    } else {
        return handleTextBasedRewrite(mesId, swipeId, option);
    }
}

async function handleChatCompletionRewrite(mesId, swipeId, option) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    if (!mesDiv) return; // Exit if we can't find the message div

    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = getSelectedTextInfo(mesId, mesDiv);

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

    // Extension streaming overrides preset streaming
    selectedPresetSettings.stream_openai = extension_settings[extensionName].useStreaming;

    if (extension_settings[extensionName].overrideMaxTokens) {
        selectedPresetSettings.openai_max_tokens = calculateTargetTokenCount(selectedRawText, option);
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

    // Get amount of words
    const wordCount = extractAllWords(selectedRawText).length;

    // Substitute {{rewrite}} macro with the selected text directly in the promptData.chat
    promptData.chat = promptData.chat.map(message => {
        if (Array.isArray(message.content)) {
            // If content is an array, process only the text entries
            message.content = message.content.map(item => {
                if (item.type === 'text') {
                    item.text = item.text.replace(/{{rewrite}}/gi, selectedRawText);
                    item.text = item.text.replace(/{{targetmessage}}/gi, fullMessage);
                    item.text = item.text.replace(/{{rewritecount}}/gi, wordCount);
                }
                return item;
            });
        } else if (typeof message.content === 'string') {
            // If content is a string, process it directly
            message.content = message.content.replace(/{{rewrite}}/gi, selectedRawText);
            message.content = message.content.replace(/{{targetmessage}}/gi, fullMessage);
            message.content = message.content.replace(/{{rewritecount}}/gi, wordCount);
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

    await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);
    getContext().activateSendButtons();
}

async function handleSimplifiedChatCompletionRewrite(mesId, swipeId, option) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    if (!mesDiv) return; // Exit if we can't find the message div

    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = getSelectedTextInfo(mesId, mesDiv);

    // Get the text completion prompt based on the option
    let promptTemplate;
    switch (option) {
        case 'Rewrite':
            promptTemplate = extension_settings[extensionName].textRewritePrompt;
            break;
        case 'Shorten':
            promptTemplate = extension_settings[extensionName].textShortenPrompt;
            break;
        case 'Expand':
            promptTemplate = extension_settings[extensionName].textExpandPrompt;
            break;
        default:
            return; // Exit if the option is not recognized
    }

    // Get amount of words
    const wordCount = extractAllWords(selectedRawText).length;

    // Replace macros in the prompt template
    let prompt = getContext().substituteParams(promptTemplate);

    prompt = prompt
        .replace(/{{rewrite}}/gi, selectedRawText)
        .replace(/{{targetmessage}}/gi, fullMessage)
        .replace(/{{rewritecount}}/gi, wordCount);

    // Create a simplified chat format
    const simplifiedChat = [
        {
            role: "system",
            content: prompt
        }
    ];

    // Create a new AbortController
    abortController = new AbortController();

    // Store the necessary data in the signal
    abortController.signal.mesDiv = mesDiv;
    abortController.signal.mesId = mesId;
    abortController.signal.swipeId = swipeId;
    abortController.signal.highlightDuration = extension_settings[extensionName].highlightDuration;

    // Show the stop button
    getContext().deactivateSendButtons();

    const res = await sendOpenAIRequest('normal', simplifiedChat, abortController.signal);
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
        newText = res?.choices?.[0]?.message?.content ?? '';
        const highlightedNewText = document.createElement('span');
        highlightedNewText.className = 'animated-highlight';
        highlightedNewText.textContent = newText;

        range.deleteContents();
        range.insertNode(highlightedNewText);
    }

    // Remove highlight after x seconds when streaming is complete
    const highlightDuration = extension_settings[extensionName].highlightDuration;
    setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

    await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);
    getContext().activateSendButtons();
}

async function handleTextBasedRewrite(mesId, swipeId, option) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    if (!mesDiv) return; // Exit if we can't find the message div

    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = getSelectedTextInfo(mesId, mesDiv);

    // Get the selected model and option-specific prompt
    const selectedModel = extension_settings[extensionName].selectedModel;
    let promptTemplate;
    switch (option) {
        case 'Rewrite':
            promptTemplate = extension_settings[extensionName].textRewritePrompt;
            break;
        case 'Shorten':
            promptTemplate = extension_settings[extensionName].textShortenPrompt;
            break;
        case 'Expand':
            promptTemplate = extension_settings[extensionName].textExpandPrompt;
            break;
        default:
            console.error('Unknown rewrite option:', option);
            return;
    }

    // Get amount of words
    const wordCount = extractAllWords(selectedRawText).length;

    // Replace macros in the prompt template
    let prompt = getContext().substituteParams(promptTemplate);

    prompt = prompt
        .replace(/{{rewrite}}/gi, selectedRawText)
        .replace(/{{targetmessage}}/gi, fullMessage)
        .replace(/{{rewritecount}}/gi, wordCount);

    let generateData;
    let amount_gen;

    if (extension_settings[extensionName].useDynamicTokens) {
        amount_gen = calculateTargetTokenCount(selectedRawText, option);
    } else {
        switch (option) {
            case 'Rewrite':
                amount_gen = extension_settings[extensionName].rewriteTokens;
                break;
            case 'Shorten':
                amount_gen = extension_settings[extensionName].shortenTokens;
                break;
            case 'Expand':
                amount_gen = extension_settings[extensionName].expandTokens;
                break;
        }
    }

    // Prepare generation data based on the selected model
    switch (main_api) {
        case 'novel':
            const novelSettings = novelai_settings[novelai_setting_names[nai_settings.preset_settings_novel]];
            generateData = getNovelGenerationData(prompt, novelSettings, amount_gen, false, false, null, 'quiet');
            break;
        case 'textgenerationwebui':
            generateData = getTextGenGenerationData(prompt, amount_gen, false, false, null, 'quiet');
            break;
        // Add more cases for other text-based models as needed
        default:
            console.error('Unsupported model:', selectedModel);
            return;
    }

    // Create a new AbortController
    abortController = new AbortController();

    // Store the necessary data in the signal
    abortController.signal.mesDiv = mesDiv;
    abortController.signal.mesId = mesId;
    abortController.signal.swipeId = swipeId;
    abortController.signal.highlightDuration = extension_settings[extensionName].highlightDuration;

    // Show the stop button
    getContext().deactivateSendButtons();
    let res;
    if (extension_settings[extensionName].useStreaming) {
        switch (main_api) {
            case 'textgenerationwebui':
                res = await generateTextGenWithStreaming(generateData, abortController.signal);
                break;
            case 'novel':
                res = await generateNovelWithStreaming(generateData, abortController.signal);
                break;
            default:
                throw new Error('Streaming is enabled, but the current API does not support streaming.');
        }
    } else {
        // Shamelessly copied from script.js
        function getGenerateUrl(api) {
            switch (api) {
                case 'textgenerationwebui':
                    return '/api/backends/text-completions/generate';
                case 'novel':
                    return '/api/novelai/generate';
                default:
                    throw new Error(`Unknown API: ${api}`);
            }
        }

        const response = await fetch(getGenerateUrl(main_api), {
            method: 'POST',
            headers: getRequestHeaders(),
            cache: 'no-cache',
            body: JSON.stringify(generateData),
            signal: abortController.signal,
        });

        if (!response.ok) {
            const error = await response.json();
            throw error;
        }

        res = await response.json();
    }

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
        if (main_api === 'novel') newText = res.output;
        const highlightedNewText = document.createElement('span');
        highlightedNewText.className = 'animated-highlight';
        highlightedNewText.textContent = newText;

        range.deleteContents();
        range.insertNode(highlightedNewText);
    }

    // Remove highlight after x seconds when streaming is complete
    const highlightDuration = extension_settings[extensionName].highlightDuration;
    setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

    await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);
    getContext().activateSendButtons();
}

function calculateTargetTokenCount(selectedText, option) {
    const baseTokenCount = getTokenCount(selectedText);
    const useDynamicTokens = extension_settings[extensionName].useDynamicTokens;
    const dynamicTokenMode = extension_settings[extensionName].dynamicTokenMode;
    let result;

    if (useDynamicTokens) {
        if (dynamicTokenMode === 'additive') {
            let modifier;
            switch (option) {
                case 'Rewrite':
                    modifier = extension_settings[extensionName].rewriteTokensAdd;
                    break;
                case 'Shorten':
                    modifier = extension_settings[extensionName].shortenTokensAdd;
                    break;
                case 'Expand':
                    modifier = extension_settings[extensionName].expandTokensAdd;
                    break;
            }
            result = baseTokenCount + modifier;
        } else { // multiplicative
            let multiplier;
            switch (option) {
                case 'Rewrite':
                    multiplier = extension_settings[extensionName].rewriteTokensMult;
                    break;
                case 'Shorten':
                    multiplier = extension_settings[extensionName].shortenTokensMult;
                    break;
                case 'Expand':
                    multiplier = extension_settings[extensionName].expandTokensMult;
                    break;
            }
            result = baseTokenCount * multiplier;
        }
    } else {
        switch (option) {
            case 'Rewrite':
                result = extension_settings[extensionName].rewriteTokens;
                break;
            case 'Shorten':
                result = extension_settings[extensionName].shortenTokens;
                break;
            case 'Expand':
                result = extension_settings[extensionName].expandTokens;
                break;
        }
    }

    return Math.max(1, Math.round(result)); // Ensure at least 1 token and round to nearest integer
}

async function handleUndo(event) {
    const mesId = event.target.dataset.mesId;
    const change = changeHistory.findLast(change => change.mesId === mesId);

    if (change) {
        const context = getContext();
        const messageDiv = document.querySelector(`[mesid="${mesId}"]`);

        if (!messageDiv || !context.chat[mesId]) {
            console.error('Message not found for undo operation');
            return;
        }

        // Update the chat context
        context.chat[mesId].mes = change.originalContent;

        // Only update swipes if they exist
        if (change.swipeId !== undefined && context.chat[mesId].swipes) {
            context.chat[mesId].swipes[change.swipeId] = change.originalContent;
        }

        // Update the UI
        const mesTextElement = messageDiv.querySelector('.mes_text');
        if (mesTextElement) {
            mesTextElement.innerHTML = messageFormatting(
                change.originalContent,
                context.name2,
                context.chat[mesId].isSystem,
                context.chat[mesId].isUser,
                mesId
            );
            addCopyToCodeBlocks(mesTextElement);
        }

        // Save the chat
        await context.saveChat();

        // Remove this change from history
        changeHistory = changeHistory.filter(c => c !== change);

        // Update undo buttons
        updateUndoButtons();
    }
}

async function saveRewrittenText(mesId, swipeId, fullMessage, startOffset, endOffset, newText) {
    const context = getContext();

    // Get the prefix and suffix to remove from the settings
    const removePrefix = extension_settings[extensionName].removePrefix || '';
    const removeSuffix = extension_settings[extensionName].removeSuffix || '';

    // Remove prefix if present
    if (removePrefix && newText.startsWith(removePrefix)) {
        newText = newText.slice(removePrefix.length);
    }

    // Remove suffix if present
    if (removeSuffix && newText.endsWith(removeSuffix)) {
        newText = newText.slice(0, -removeSuffix.length);
    }

    // Create the new message with the rewritten section
    const newMessage =
        fullMessage.substring(0, startOffset) +
        newText +
        fullMessage.substring(endOffset);

    // Save the change to the history
    saveLastChange(mesId, swipeId, fullMessage, newMessage);

    // Update the main message
    context.chat[mesId].mes = newMessage;

    // Update the swipe if it exists
    if (swipeId !== undefined && context.chat[mesId].swipes && context.chat[mesId].swipes[swipeId]) {
        context.chat[mesId].swipes[swipeId] = newMessage;
    }

    // Save and reload the chat
    await context.saveChat();
}
