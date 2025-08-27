import { sendOpenAIRequest, oai_settings } from "../../../openai.js";
import { extractAllWords } from "../../../utils.js";
import { getTokenCount } from "../../../tokenizers.js";
import { getNovelGenerationData, generateNovelWithStreaming, nai_settings } from "../../../nai-settings.js";
import { generateHorde, MIN_LENGTH } from "../../../horde.js";
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
    generateRaw,
} from "../../../../script.js";
import { extension_settings, getContext } from "../../../extensions.js";
import { getRegexedString, regex_placement } from '../../regex/engine.js'; // Import from built-in regex extension

const extensionName = "rewrite-extension";
const extensionFolderPath = `scripts/extensions/third-party/${extensionName}`;

const undo_steps = 15;

// Default settings
const defaultSettings = {
    rewritePreset: "",
    shortenPreset: "",
    expandPreset: "",
    instructPreset: "", 
    custom1Preset: "",
    custom2Preset: "",
    custom3Preset: "",
    custom4Preset: "",
    custom5Preset: "",
    rewriteButtonName: "Rewrite",
    shortenButtonName: "Shorten",
    expandButtonName: "Expand",
    instructButtonName: "Instruct",
    custom1ButtonName: "Custom1",
    custom2ButtonName: "Custom2",
    custom3ButtonName: "Custom3",
    custom4ButtonName: "Custom4",
    custom5ButtonName: "Custom5",
    highlightDuration: 3000,
    selectedModel: "chat_completion",
    textRewritePrompt: `[INST]Rewrite this section of text: """{{rewrite}}""" while keeping the same content, general style and length. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the rewritten text without any comments: `,
    textShortenPrompt: `[INST]Rewrite this section of text: """{{rewrite}}""" while keeping the same content, general style. Do not list alternatives and only print the result without prefix or suffix. Shorten it by roughly 20%.[/INST]

Sure, here is only the rewritten text without any comments: `,
    textExpandPrompt: `[INST]Rewrite this section of text: """{{rewrite}}""" while keeping the same content, general style. Do not list alternatives and only print the result without prefix or suffix. Lengthen it by roughly 20%.[/INST]

Sure, here is only the rewritten text without any comments: `,
    textInstructPrompt: `[INST]Rewrite this section of text: """{{rewrite}}""" according to the following instructions: "{{instruct_instructions}}". Keep the general style. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the rewritten text without any comments: `, 
    textCustom1Prompt: `[INST]Improve the style and clarity of this text: """{{rewrite}}""" while keeping the same content and general length. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the improved text without any comments: `,
    textCustom2Prompt: `[INST]Adjust the tone of this text: """{{rewrite}}""" to be more professional while keeping the same content and general length. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the adjusted text without any comments: `,
    textCustom3Prompt: `[INST]Enhance the clarity and readability极 of this text: """{{rewrite}}""" while keeping the same content and general length. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the enhanced text without any comments: `,
    textCustom4Prompt: `[INST]Adjust the formality level of this text: """{{rewrite}}""" while keeping the same content and general length. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the adjusted text without any comments: `,
    textCustom5Prompt: `[INST]Add more creativity and engagement to this text: """{{rewrite}}""" while keeping the same core content and general length. Do not list alternatives and only print the result without prefix or suffix.[/INST]

Sure, here is only the enhanced text without any comments: `,
    useStreaming: true,
    useDynamicTokens: true,
    dynamicTokenMode: 'multiplicative',
    rewriteTokens: 100,
    shortenTokens: 50,
    expandTokens: 150,
    instructTokens: 100, 
    custom1Tokens: 100,
    custom2Tokens: 100,
    custom3Tokens: 100,
    custom4Tokens: 100,
    custom5Tokens: 100,
    rewriteTokensAdd: 0,
    shortenTokensAdd: -50,
    expandTokensAdd: 50,
    instructTokensAdd: 0, 
    custom1TokensAdd: 0,
    custom2TokensAdd: 0,
    custom3TokensAdd: 0,
    custom4TokensAdd: 0,
    custom5TokensAdd: 0,
    rewriteTokensMult: 1.05,
    shortenTokensMult: 0.8,
    expandTokensMult: 1.5,
    instructTokensMult: 1.0, 
    custom1TokensMult: 1.0,
    custom2TokensMult: 1.0,
    custom3TokensMult: 1.0,
    custom4TokensMult: 1.0,
    custom5TokensMult: 1.0,
    removePrefix: `"`,
    removeSuffix: `"`,
    overrideMaxTokens: true,
    showRewrite: true,
    showShorten: true,
    showExpand: true,
    showInstruct: true, 
    showCustom1: true,
    showCustom2: true,
    showCustom3: true,
    showCustom4: true,
    showCustom5: true,
    showDelete: true,
    applyRegexOnRewrite: true, // New setting to control regex application
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
    $("#instruct_preset").val(getSetting('instructPreset', defaultSettings.instructPreset)); 
    $("#custom1_preset").val(getSetting('custom1Preset', defaultSettings.custom1Preset));
    $("#custom2_preset").val(getSetting('custom2Preset', defaultSettings.custom2Preset));
    $("#custom3_preset").val(getSetting('custom3Preset', defaultSettings.custom3Preset));
    $("#custom4_preset").val(getSetting('custom4Preset', defaultSettings.custom4Preset));
    $("#custom5_preset").val(getSetting('custom5Preset', defaultSettings.custom5Preset));
    $("#highlight_duration").val(getSetting('highlightDuration', defaultSettings.highlightDuration));
    $("#rewrite_extension_model_select").val(getSetting('selectedModel', defaultSettings.selectedModel));
    $("#text_rewrite_prompt").val(getSetting('textRewritePrompt', defaultSettings.textRewritePrompt));
    $("#text_shorten_prompt").val(getSetting('textShortenPrompt', defaultSettings.textShortenPrompt));
    $("#text_expand_prompt").val(getSetting('textExpandPrompt', defaultSettings.textExpandPrompt));
    $("#text_instruct_prompt").val(getSetting('textInstructPrompt', defaultSettings.textInstructPrompt)); 
    $("#text_custom1_prompt").val(getSetting('textCustom1Prompt', defaultSettings.textCustom1Prompt));
    $("#text_custom2_prompt").val(getSetting('textCustom2Prompt', defaultSettings.textCustom2Prompt));
    $("#text_custom3_prompt").val(getSetting('textCustom3Prompt', defaultSettings.textCustom3Prompt));
    $("#text_custom4_prompt").val(getSetting('textCustom4Prompt', defaultSettings.textCustom4Prompt));
    $("#text_custom5_prompt").val(getSetting('textCustom5Prompt', defaultSettings.textCustom5Prompt));
    $("#use_streaming").prop('checked', getSetting('useStreaming', defaultSettings.useStreaming));
    $("#use_dynamic_tokens").prop('checked', getSetting('useDynamicTokens', defaultSettings.useDynamicTokens));
    $("#dynamic_token_mode").val(getSetting('dynamicTokenMode', defaultSettings.dynamicTokenMode));
    $("#rewrite_tokens").val(getSetting('rewriteTokens', defaultSettings.rewriteTokens));
    $("#shorten_tokens").val(getSetting('shortenTokens', defaultSettings.shortenTokens));
    $("#expand_tokens").val(getSetting('expandTokens', defaultSettings.expandTokens));
    $("#instruct_tokens").val(getSetting('instructTokens', defaultSettings.instructTokens)); 
    $("#custom1_tokens").val(getSetting('custom1Tokens', defaultSettings.custom1Tokens));
    $("#custom2_tokens").val(getSetting('custom2Tokens', defaultSettings.custom2Tokens));
    $("#custom3_tokens").val(getSetting('custom3Tokens', defaultSettings.custom3Tokens));
    $("#custom4_tokens").val(getSetting('custom4Tokens', defaultSettings.custom4Tokens));
    $("#custom5_tokens").val(getSetting('custom5Tokens', defaultSettings.custom5Tokens));
    $("#rewrite_tokens_add").val(getSetting('rewriteTokensAdd', defaultSettings.rewriteTokensAdd));
    $("#shorten_tokens_add").val(getSetting('shortenTokensAdd', defaultSettings.shortenTokensAdd));
    $("#expand_tokens_add").val(getSetting('expandTokensAdd', defaultSettings.expandTokensAdd));
    $("#instruct_tokens_add").val(getSetting('instructTokensAdd', defaultSettings.instructTokensAdd)); 
    $("#custom1_tokens_add").val(getSetting('custom1TokensAdd', defaultSettings.custom1TokensAdd));
    $("#custom2_tokens_add").val(getSetting('custom2TokensAdd', defaultSettings.custom2TokensAdd));
    $("#custom3_tokens_add").val(getSetting('custom3TokensAdd', defaultSettings.custom3TokensAdd));
    $("#custom4_tokens_add").val(getSetting('custom4TokensAdd', defaultSettings.custom4TokensAdd));
    $("#custom5_tokens_add").val(getSetting('custom5TokensAdd', defaultSettings.custom5TokensAdd));
    $("#rewrite_tokens_mult").val(getSetting('rewriteTokensMult', defaultSettings.rewriteTokensMult));
    $("#shorten_tokens_mult").val(getSetting('shortenTokensMult', defaultSettings.shortenTokensMult));
    $("#expand_tokens_mult").val(getSetting('expandTokensMult', defaultSettings.expandTokensMult));
    $("#instruct_tokens_mult").val(getSetting('instructTokensMult', defaultSettings.instructTokensMult)); 
    $("#custom1_tokens_mult").val(getSetting('custom1TokensMult', defaultSettings.custom1TokensMult));
    $("#custom2_tokens_mult").val(getSetting('custom2TokensMult', defaultSettings.custom2TokensMult));
    $("#custom3_tokens_mult").val(getSetting('custom3TokensMult', defaultSettings.custom3TokensMult));
    $("#custom4_tokens_mult").val(getSetting('custom4TokensMult', defaultSettings.custom4TokensMult));
    $("#custom5_tokens_mult").val(getSetting('custom5TokensMult', defaultSettings.custom5TokensMult));
    $("#remove_prefix").val(getSetting('removePrefix', defaultSettings.removePrefix));
    $("#remove_suffix").val(getSetting('removeSuffix', defaultSettings.removeSuffix));
    $("#override_max_tokens").prop('checked', getSetting('overrideMaxTokens', defaultSettings.overrideMaxTokens));
    $("#show_rewrite").prop('checked', getSetting('showRewrite', defaultSettings.showRewrite));
    $("#show_shorten").prop('checked', getSetting('showShorten', defaultSettings.showShorten));
    $("#show_expand").prop('checked', getSetting('showExpand', defaultSettings.showExpand));
    $("#show_instruct").prop('checked', getSetting('showInstruct', defaultSettings.showInstruct)); 
    $("#show_custom1").prop('checked', getSetting('showCustom1', defaultSettings.showCustom1));
    $("#show_custom2").prop('checked', getSetting('showCustom2', defaultSettings.showCustom2));
    $("#show_custom3").prop('checked', getSetting('showCustom3', defaultSettings.showCustom3));
    $("#show_custom4").prop('checked', getSetting('showCustom4', defaultSettings.showCustom4));
    $("#show_custom5").prop('checked', getSetting('showCustom5', defaultSettings.showCustom5));
    $("#show_delete").prop('checked', getSetting('showDelete', defaultSettings.showDelete));
    $("#apply_regex_on_rewrite").prop('checked', getSetting('applyRegexOnRewrite', defaultSettings.applyRegexOnRewrite)); // Load new setting

    // Load button name settings
    $("#rewrite_button_name").val(getSetting('rewriteButtonName', defaultSettings.rewriteButtonName));
    $("#shorten_button_name").val(getSetting('shortenButtonName', defaultSettings.shortenButtonName));
    $("#expand_button_name").val(getSetting('expandButtonName', defaultSettings.expandButtonName));
    $("#instruct_button_name").val(getSetting('instructButtonName', defaultSettings.instructButtonName));
    $("#custom1_button_name").val(getSetting('custom1ButtonName', defaultSettings.custom1ButtonName));
    $("#custom2_button_name").val(getSetting('custom2ButtonName', defaultSettings.custom2ButtonName));
    $("#custom3_button_name").val(getSetting('custom3ButtonName', defaultSettings.custom3ButtonName));
    $("#custom4_button_name").val(getSetting('custom4ButtonName', defaultSettings.custom4ButtonName));
    $("#custom5_button_name").val(getSetting('custom5ButtonName', defaultSettings.custom5ButtonName));

    // Update the UI based on loaded settings
    updateModelSettings();
    updateTokenSettings();
}

function saveSettings() {
    extension_settings[extensionName] = {
        rewritePreset: $("#rewrite_preset").val(),
        shortenPreset: $("#shorten_preset").val(),
        expandPreset: $("#expand_preset").val(),
        instructPreset: $("#instruct_preset").val(), 
        custom1Preset: $("#custom1_preset").val(),
        custom2Preset: $("#custom2_preset").val(),
        custom3Preset: $("#custom3_preset").val(),
        custom4Preset: $("#custom4_preset").val(),
        custom5Preset: $("#custom5_preset").val(),
        rewriteButtonName: $("#rewrite_button_name").val(),
        shortenButtonName: $("#shorten_button_name").val(),
        expandButtonName: $("#expand_button_name").val(),
        instructButtonName: $("#instruct_button_name").val(),
        custom1ButtonName: $("#custom1_button_name").val(),
        custom2ButtonName: $("#custom2_button_name").val(),
        custom3ButtonName: $("#custom3_button_name").val(),
        custom4ButtonName: $("#custom4_button_name").val(),
        custom5ButtonName: $("#custom5_button_name").val(),
        highlightDuration: parseInt($("#highlight_duration").val()),
        selectedModel: $("#rewrite_extension_model_select").val(),
        textRewritePrompt: $("#text_rewrite_prompt").val(),
        textShortenPrompt: $("#text_shorten_prompt").val(),
        textExpandPrompt: $("#text_expand_prompt").val(),
        textInstructPrompt: $("#text_instruct_prompt").val(), 
        textCustom1Prompt: $("#text_custom1_prompt").val(),
        textCustom2Prompt: $("#text_custom2_prompt").val(),
        textCustom3Prompt: $("#text_custom3_prompt").val(),
        textCustom4Prompt: $("#text_custom4_prompt").val(),
        textCustom5Prompt: $("#text_custom5_prompt").val(),
        useStreaming: $("#use_streaming").is(':checked'),
        useDynamicTokens: $("#use_dynamic_tokens").is(':checked'),
        dynamicTokenMode: $("#dynamic_token_mode").val(),
        rewriteTokens: parseInt($("#rewrite_tokens").val()),
        shortenTokens: parseInt($("#shorten_tokens").val()),
        expandTokens: parseInt($("#expand_tokens").val()),
        instructTokens: parseInt($("#instruct_tokens").val()), 
        custom1Tokens: parseInt($("#custom1_tokens").val()),
        custom2Tokens: parseInt($("#custom2_tokens").val()),
        custom3Tokens: parseInt($("#custom3_tokens").val()),
        custom4Tokens: parseInt($("#custom4_tokens").val()),
        custom5Tokens: parseInt($("#custom5_tokens").val()),
        rewriteTokensAdd: parseInt($("#rewrite_tokens_add").val()),
        shortenTokensAdd: parseInt($("#shorten_tokens_add").val()),
        expandTokensAdd: parseInt($("#expand_tokens_add").val()),
        instructTokensAdd: parseInt($("#instruct_tokens_add").val()), 
        custom1TokensAdd: parseInt($("#custom1_tokens_add").val()),
        custom2TokensAdd: parseInt($("#custom2_tokens_add").val()),
        custom3TokensAdd: parseInt($("#custom3_tokens_add").val()),
        custom4TokensAdd: parseInt($("#custom4_tokens_add").val()),
        custom5TokensAdd: parseInt($("#custom5_tokens_add").val()),
        rewriteTokensMult: parseFloat($("#rewrite_tokens_mult").val()),
        shortenTokensMult: parseFloat($("#shorten_tokens_mult").val()),
        expandTokensMult: parseFloat($("#expand_tokens_mult").val()),
        instructTokensMult: parseFloat($("#instruct_tokens_mult").val()), 
        custom1TokensMult: parseFloat($("#custom1_tokens_mult").val()),
        custom2TokensMult: parseFloat($("#custom2_tokens_mult").val()),
        custom3TokensMult: parseFloat($("#custom3_tokens_mult").val()),
        custom4TokensMult: parseFloat($("#custom4_tokens_mult").val()),
        custom5TokensMult: parseFloat($("#custom5_tokens_mult").val()),
        removePrefix: $("#remove_prefix").val(),
        removeSuffix: $("#remove_suffix").val(),
        overrideMaxTokens: $("#override_max_tokens").is(':checked'),
        showRewrite: $("#show_rewrite").is(':checked'),
        showShorten: $("#show_shorten").is(':checked'),
        showExpand: $("#show_expand").is(':checked'),
        showInstruct: $("#show_instruct").is(':checked'), 
        showCustom1: $("#show_custom1").is(':checked'),
        showCustom2: $("#show_custom2").is(':checked'),
        showCustom3: $("#show_custom3").is(':checked'),
        showCustom4: $("#show_custom4").is(':checked'),
        showCustom5: $("#show_custom5").is(':checked'),
        showDelete: $("#show_delete").is(':checked'),
        applyRegexOnRewrite: $("#apply_regex_on_rewrite").is(':checked'), // Save new setting
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
    const dropdowns = ['rewrite_preset', 'shorten_preset', 'expand_preset', 'instruct_preset', 'custom1_preset', 'custom2_preset', 'custom3_preset', 'custom4_preset', 'custom5_preset']; // Added custom presets
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
    $(".rewrite-extension-settings select, #highlight_duration, #text_rewrite_prompt, #text_shorten_prompt, #text_expand_prompt, #text_instruct_prompt, #text_custom1_prompt, #text_custom2_prompt, #text_custom3_prompt, #text_custom4_prompt, #text_custom5_prompt").on("change", saveSettings); // Added custom prompts
    $("#use_streaming").on("change", saveSettings);
    $("#use_dynamic_tokens, #dynamic_token_mode").on("change", () => {
        updateTokenSettings();
        saveSettings();
    });
    $("#rewrite_tokens, #shorten_tokens, #expand_tokens, #instruct_tokens, #custom1_tokens, #custom2_tokens, #custom3_tokens, #custom4_tokens, #custom5_tokens, #rewrite_tokens_add, #shorten_tokens_add, #expand_tokens_add, #instruct_tokens_add, #custom1_tokens_add, #custom2_tokens_add, #custom3_tokens_add, #custom4_tokens_add, #custom5_tokens_add, #rewrite_tokens_mult, #shorten_tokens_mult, #expand_tokens_mult, #instruct_tokens_mult, #custom1_tokens_mult, #custom2_tokens_mult, #custom3_tokens_mult, #custom4_tokens_mult, #custom5_tokens_mult").on("input", saveSettings); // Added custom token inputs
    $("#remove_prefix, #remove_suffix").on("change", saveSettings);
    $("#override_max_tokens").on("change", saveSettings);
    $("#show_rewrite, #show_shorten, #show_expand, #show_instruct, #show_custom1, #show_custom2, #show_custom3, #show_custom4, #show_custom5, #show_delete").on("change", saveSettings); // Added custom visibility checkboxes
    $("#apply_regex_on_rewrite").on("change", saveSettings); // Add listener for new checkbox

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

async function getInstructInstructionsFromPopup() {
    const { callPopup } = getContext();
    try {
        const instructions = await callPopup('Enter rewrite instructions:', 'input');

        // Introduce a zero-delay setTimeout to yield to the event loop
        await new Promise(resolve => setTimeout(resolve, 0));

        return instructions;
    } catch (error) {
        console.error("[Rewrite Extension] Error during custom instruction popup:", error);
        return null;
    } finally {
    }
}

async function handleMenuItemClick(e) {
    e.preventDefault();
    e.stopPropagation();

    const option = e.target.dataset.option;
    const selection = window.getSelection();

    // Ensure there's a selection and a range
    if (!selection || selection.rangeCount === 0) {
        removeRewriteMenu();
        return;
    }

    // Capture the range *before* any awaits or potential selection changes
    const initialRange = selection.getRangeAt(0).cloneRange();
    const selectedText = initialRange.toString().trim();

    if (selectedText) {
        const mesTextElement = findClosestMesText(selection.anchorNode);
        if (mesTextElement) {
            const messageDiv = findMessageDiv(mesTextElement);
            if (messageDiv) {
                const mesId = messageDiv.getAttribute('mesid');
                const swipeId = messageDiv.getAttribute('swipeid');

                if (option === 'Delete') {
                    // Pass the initially captured range to handleDeleteSelection
                    await handleDeleteSelection(mesId, swipeId, initialRange);
                } else if (option === 'Instruct') {
                    const instructInstructions = await getInstructInstructionsFromPopup();
                    if (instructInstructions !== null && instructInstructions.trim() !== '') { // Proceed only if user entered text and didn't cancel
                        // Get selectionInfo *after* await and *before* handleRewrite
                        // Pass the initially captured range
                        const selectionInfo = getSelectedTextInfo(mesId, mesTextElement, initialRange);
                        if (!selectionInfo) {
                             console.error("[Rewrite Extension] Failed to get selectionInfo for Instruct rewrite!");
                             return; // Prevent calling with undefined
                        }
                        await handleRewrite(mesId, swipeId, option, instructInstructions, selectionInfo); // Use the locally scoped selectionInfo
                    } else {
                        // User cancelled or entered empty instructions
                    }
                } else {
                    // For other rewrite options, get selectionInfo right before the call
                    // Pass the initially captured range
                    const selectionInfo = getSelectedTextInfo(mesId, mesTextElement, initialRange); // Get selectionInfo here
                    if (!selectionInfo) {
                         console.error(`[Rewrite Extension] Failed to get selectionInfo for ${option} rewrite!`);
                         return; // Prevent calling with undefined
                    }
                    await handleRewrite(mesId, swipeId, option, null, selectionInfo); // Use the locally scoped selectionInfo
                }
            }
        }
    }

    removeRewriteMenu();
    window.getSelection().removeAllRanges();
}

// Modify signature to accept the captured range
async function handleDeleteSelection(mesId, swipeId, range) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    // Use the passed-in range to get selection info
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset } = getSelectedTextInfo(mesId, mesDiv, range);

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
        { name: extension_settings[extensionName].rewriteButtonName || 'Rewrite', show: extension_settings[extensionName].showRewrite, id: 'Rewrite' },
        { name: extension_settings[extensionName].shortenButtonName || 'Shorten', show: extension_settings[extensionName].showShorten, id: 'Shorten' },
        { name: extension_settings[extensionName].expandButtonName || 'Expand', show: extension_settings[extensionName].showExpand, id: 'Expand' },
        { name: extension_settings[extensionName].instructButtonName || 'Instruct', show: extension_settings[extensionName].showInstruct, id: 'Instruct' }, 
        { name: extension_settings[extensionName].custom1ButtonName || 'Custom1', show: extension_settings[extensionName].showCustom1, id: 'Custom1' },
        { name: extension_settings[extensionName].custom2ButtonName || 'Custom2', show: extension_settings[extensionName].showCustom2, id: 'Custom2' },
        { name: extension_settings[extensionName].custom3ButtonName || 'Custom3', show: extension_settings[extensionName].showCustom3, id: 'Custom3' },
        { name: extension_settings[extensionName].custom4ButtonName || 'Custom4', show: extension_settings[extensionName].showCustom4, id: 'Custom4' },
        { name: extension_settings[extensionName].custom5ButtonName || 'Custom5', show: extension_settings[extensionName].showCustom5, id: 'Custom5' },
        { name: 'Delete', show: extension_settings[extensionName].showDelete, id: 'Delete' }
    ];
    options.forEach(option => {
        if (option.show) {
            let li = document.createElement('li');
            li.className = 'list-group-item ctx-item';
            li.textContent = option.name;
            li.addEventListener('mousedown', handleMenuItemClick);
            li.addEventListener('touchstart', handleMenuItemClick);
            li.dataset.option = option.id;
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

// Modify signature to accept the captured range
function getSelectedTextInfo(mesId, mesDiv, range) {
    // Removed: const selection = window.getSelection();
    // Removed: const range = selection.getRangeAt(0); - Use the passed-in range directly

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
    let rawStartOffset = mapping.formattedToRaw(startOffset);
    let rawEndOffset = mapping.formattedToRaw(endOffset);

    // Heuristic: Adjust offsets to include surrounding markdown if selection seems to abut it
    // Check for italics (*)
    if (rawStartOffset > 0 && rawEndOffset < fullMessage.length &&
        fullMessage[rawStartOffset - 1] === '*' && fullMessage[rawEndOffset] === '*') {
        // Avoid expanding if it looks like bold/bold-italics boundary
        const prevChar = rawStartOffset > 1 ? fullMessage[rawStartOffset - 2] : null;
        const nextChar = rawEndOffset + 1 < fullMessage.length ? fullMessage[rawEndOffset + 1] : null;
        if (prevChar !== '*' && nextChar !== '*') {
            rawStartOffset--;
            rawEndOffset++;
        }
    }
    // Check for bold (**) - ensure we don't double-adjust if italics check already expanded
    else if (rawStartOffset > 1 && rawEndOffset < fullMessage.length - 1 &&
             fullMessage.substring(rawStartOffset - 2, rawStartOffset) === '**' &&
             fullMessage.substring(rawEndOffset, rawEndOffset + 2) === '**') {
        // Avoid expanding if it looks like bold-italics boundary
        const prevChar = rawStartOffset > 2 ? fullMessage[rawStartOffset - 3] : null;
        const nextChar = rawEndOffset + 2 < fullMessage.length ? fullMessage[rawEndOffset + 2] : null;
        if (prevChar !== '*' && nextChar !== '*') {
            rawStartOffset -= 2;
            rawEndOffset += 2;
        }
    }
    // Note: This doesn't handle ***bold italics*** or nested cases perfectly, but covers common scenarios.

    // Get the selected raw text using potentially adjusted offsets
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

// Updated handleRewrite signature to accept selectionInfo
async function handleRewrite(mesId, swipeId, option, customInstructions = null, selectionInfo) {
    if (!selectionInfo) {
        console.error("[Rewrite Extension] handleRewrite called without selectionInfo!");
        return; // Cannot proceed without selection info
    }

    if (main_api === 'openai') {
        const selectedModel = extension_settings[extensionName].selectedModel;
        if (selectedModel === 'chat_completion') {
            return handleChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo); // Pass selectionInfo
        } else {
            return handleSimplifiedChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo); // Pass selectionInfo
        }
    } else {
        return handleTextBasedRewrite(mesId, swipeId, option, customInstructions, selectionInfo); // Pass selectionInfo
    }
}

// Updated signature to accept selectionInfo
async function handleChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo) {
    // Use pre-captured selection info
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = selectionInfo;
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`); // Keep getting mesDiv for highlight/DOM ops
    if (!mesDiv) { // Add check for mesDiv existence
        console.error("[Rewrite Extension] Could not find mesDiv in handleChatCompletionRewrite.");
        return;
    }

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
        case 'Instruct': // New case
            selectedPreset = extension_settings[extensionName].instructPreset;
            break;
        case 'Custom1':
            selectedPreset = extension_settings[extensionName].custom1Preset;
            break;
        case 'Custom2':
            selectedPreset = extension_settings[extensionName].custom2Preset;
            break;
        case 'Custom3':
            selectedPreset = extension_settings[extensionName].custom3Preset;
            break;
        case 'Custom4':
            selectedPreset = extension_settings[extensionName].custom4Preset;
            break;
        case 'Custom5':
            selectedPreset = extension_settings[extensionName].custom5Preset;
            break;
        default:
            console.error("Unknown rewrite option:", option);
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

    // Always generate the base prompt using the selected preset
    const promptReadyPromise = new Promise(resolve => {
        eventSource.once(event_types.CHAT_COMPLETION_PROMPT_READY, resolve);
    });
    getContext().generate('normal', {}, true); // Trigger prompt generation
    const promptData = await promptReadyPromise; // Wait for the generated prompt
    let chatToSend = promptData.chat; // Start with the generated chat array

    // Inject custom instructions if applicable
    if (option === 'Instruct' && customInstructions) {
        // Find the last user message to append to
        let targetMessageIndex = -1;
        for (let i = chatToSend.length - 1; i >= 0; i--) {
            if (chatToSend[i].role === 'user') {
                targetMessageIndex = i;
                break;
            }
        }

        if (targetMessageIndex !== -1) {
            const targetMessage = chatToSend[targetMessageIndex];
            const instructionText = `\n\nAdditional Instructions:\n${customInstructions}`;

            if (Array.isArray(targetMessage.content)) {
                // Find the last text part or add a new one
                let lastTextPartIndex = -1;
                for (let j = targetMessage.content.length - 1; j >= 0; j--) {
                    if (targetMessage.content[j].type === 'text') {
                        lastTextPartIndex = j;
                        break;
                    }
                }
                if (lastTextPartIndex !== -1) {
                    targetMessage.content[lastTextPartIndex].text += instructionText;
                } else {
                    // Should not happen with standard prompts, but handle just in case
                    targetMessage.content.push({ type: 'text', text: instructionText });
                }
            } else if (typeof targetMessage.content === 'string') {
                targetMessage.content += instructionText;
            }
        } else {
            console.warn('[Rewrite Extension] Could not find a user message in the generated prompt to inject custom instructions into.');
            // Optionally, could append a new user message, but might break formatting
            // chatToSend.push({ role: "user", content: `Additional Instructions:\n${customInstructions}` });
        }
    }

    // Substitute standard macros AFTER potential custom instruction injection
    const wordCount = extractAllWords(selectedRawText).length;
    chatToSend = chatToSend.map(message => {
        if (Array.isArray(message.content)) {
            message.content = message.content.map(item => {
                if (item.type === 'text') {
                    item.text = item.text.replace(/{{rewrite}}/gi, selectedRawText);
                    item.text = item.text.replace(/{{targetmessage}}/gi, fullMessage);
                    item.text = item.text.replace(/{{rewritecount}}/gi, wordCount);
                }
                return item;
            });
        } else if (typeof message.content === 'string') {
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

    let res;
    try {

        // Send the request with the prepared chat
        res = await sendOpenAIRequest('normal', chatToSend, abortController.signal);
    } catch (error) {
        console.error('[Rewrite Extension] Error during sendOpenAIRequest:', error);
        toastr.error("Rewrite failed. Check browser console (F12) for details.", "Rewrite Error");
        // Ensure cleanup happens even on error
    } finally {
        window.getSelection().removeAllRanges();
        // Restore the original settings (moved to finally)
        Object.assign(oai_settings, prev_oai_settings);
        getContext().activateSendButtons();
    }

    // If the request failed, res will be undefined, stop further processing
    if (res === undefined) {
        // Remove highlight immediately if the request failed before starting streaming/display
        removeHighlight(mesDiv, mesId, swipeId);
        return;
    }

    let newText = '';
    try {
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
        
        // Format the new text for proper display with paragraph spacing
        const messageData = getContext().chat[mesId];
        const formattedNewText = messageFormatting(
            newText,
            getContext().name2,
            messageData.isSystem,
            messageData.isUser,
            mesId
        );
        highlightedNewText.innerHTML = formattedNewText;

        range.deleteContents();
        range.insertNode(highlightedNewText);
    }

        // Remove highlight after x seconds when processing is complete
        const highlightDuration = extension_settings[extensionName].highlightDuration;
        setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

        await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);

    } catch (error) {
        console.error('[Rewrite Extension] Error processing API response:', error);
        toastr.error("Failed to process rewrite response. Check console.", "Processing Error");
        // Ensure highlight is removed if processing fails
        removeHighlight(mesDiv, mesId, swipeId);
    }
    // activateSendButtons is now handled in the finally block above
}

// Updated signature to accept selectionInfo
async function handleSimplifiedChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo) {
    // Use pre-captured selection info
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = selectionInfo;
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`); // Keep getting mesDiv for highlight/DOM ops
    if (!mesDiv) { // Add check for mesDiv existence
        console.error("[Rewrite Extension] Could not find mesDiv in handleSimplifiedChatCompletionRewrite.");
        return;
    }
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
        case 'Instruct': // New case
            promptTemplate = extension_settings[extensionName].textInstructPrompt;
            break;
        case 'Custom1':
            promptTemplate = extension_settings[extensionName].textCustom1Prompt;
            break;
        case 'Custom2':
            promptTemplate = extension_settings[extensionName].textCustom2Prompt;
            break;
        case 'Custom3':
            promptTemplate = extension_settings[extensionName].textCustom3Prompt;
            break;
        case 'Custom4':
            promptTemplate = extension_settings[extensionName].textCustom4Prompt;
            break;
        case 'Custom5':
            promptTemplate = extension_settings[extensionName].textCustom5Prompt;
            break;
        default:
            console.error("Unknown rewrite option:", option);
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

    // Inject custom instructions if applicable
    if (option === 'Instruct') {
        if (prompt.includes('{{instruct_instructions}}')) {
            prompt = prompt.replace(/{{instruct_instructions}}/gi, customInstructions);
        } else {
            // Append if macro is missing (basic fallback)
            prompt += `\n\nInstructions: ${customInstructions}`;
        }
    }

    // Create a simplified chat format
    const simplifiedChat = [
        {
            role: "system",
            content: prompt
        }
    ];

    // Save current oai_settings and override streaming setting
    const prev_oai_settings = Object.assign({}, oai_settings);
    oai_settings.stream_openai = extension_settings[extensionName].useStreaming;

    // Create a new AbortController
    abortController = new AbortController();

    // Store the necessary data in the signal
    abortController.signal.mesDiv = mesDiv;
    abortController.signal.mesId = mesId;
    abortController.signal.swipeId = swipeId;
    abortController.signal.highlightDuration = extension_settings[extensionName].highlightDuration;
    abortController.signal.prev_oai_settings = prev_oai_settings; // Store for cleanup

    // Show the stop button
    getContext().deactivateSendButtons();

    let res;
    try {
        res = await sendOpenAIRequest('normal', simplifiedChat, abortController.signal);
    } finally {
        // Restore oai_settings
        Object.assign(oai_settings, prev_oai_settings);
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
        newText = res?.choices?.[0]?.message?.content ?? '';
        const highlightedNewText = document.createElement('span');
        highlightedNewText.className = 'animated-highlight';
        
        // Format the new text for proper display with paragraph spacing
        const messageData = getContext().chat[mesId];
        const formattedNewText = messageFormatting(
            newText,
            getContext().name2,
            messageData.isSystem,
            messageData.isUser,
            mesId
        );
        highlightedNewText.innerHTML = formattedNewText;

        range.deleteContents();
        range.insertNode(highlightedNewText);
    }

    // Remove highlight after x seconds when streaming is complete
    const highlightDuration = extension_settings[extensionName].highlightDuration;
    setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

    await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);
    getContext().activateSendButtons();
}

// Updated signature to accept selectionInfo
async function handleTextBasedRewrite(mesId, swipeId, option, customInstructions, selectionInfo) {
    // Use pre-captured selection info
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = selectionInfo;
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`); // Keep getting mesDiv for highlight/DOM ops
    if (!mesDiv) { // Add check for mesDiv existence
        console.error("[Rewrite Extension] Could not find mesDiv in handleTextBasedRewrite.");
        return;
    }
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
        case 'Instruct': // New case
            promptTemplate = extension_settings[extensionName].textInstructPrompt;
            break;
        case 'Custom1':
            promptTemplate = extension_settings[extensionName].textCustom1Prompt;
            break;
        case 'Custom2':
            promptTemplate = extension_settings[extensionName].textCustom2Prompt;
            break;
        case 'Custom3':
            promptTemplate = extension_settings[extensionName].textCustom3Prompt;
            break;
        case 'Custom4':
            promptTemplate = extension_settings[extensionName].textCustom4Prompt;
            break;
        case 'Custom5':
            promptTemplate = extension_settings[extensionName].textCustom5Prompt;
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

    // Inject custom instructions if applicable
    if (option === 'Instruct') {
        if (prompt.includes('{{instruct_instructions}}')) {
            prompt = prompt.replace(/{{instruct_instructions}}/gi, customInstructions);
        } else {
            // Append if macro is missing (basic fallback)
            prompt += `\n\nInstructions: ${customInstructions}`;
        }
    }

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
            case 'Instruct': // New case
                amount_gen = extension_settings[extensionName].instructTokens;
                break;
        }
    }

    // Prepare generation data based on the selected model
    switch (main_api) {
        case 'novel':
            const novelSettings = novelai_settings[novelai_setting_names[nai_settings.preset_settings_novel]];
            generateData = getNovelGenerationData(prompt, novelSettings, amount_gen, true, false, null, 'quiet');
            break;
        case 'textgenerationwebui':
            generateData = getTextGenGenerationData(prompt, amount_gen, true, false, null, 'quiet');
            break;
        case 'koboldhorde':
            if (option === 'Instruct') {
                // For Instruct Horde, use the manually constructed prompt directly
                // We need a basic structure for generateHorde, mimicking what getContext().generate would provide
                generateData = {
                    prompt: prompt, // Use the manually constructed prompt
                    max_length: Math.max(amount_gen, MIN_LENGTH),
                    // Include other necessary default parameters if generateHorde requires them
                    // Based on generateHorde usage, 'quiet' and potentially others might be needed.
                    quiet: true, // Often used in background generation
                };
            } else {
                // Existing logic for non-custom Horde rewrites
                const promptReadyPromise = new Promise(resolve => {
                    eventSource.once(event_types.GENERATE_AFTER_DATA, resolve);
                });
                getContext().generate('normal', {}, true); // Trigger standard prompt generation
                generateData = await promptReadyPromise; // Wait for the generated data
                generateData.max_length = Math.max(amount_gen, MIN_LENGTH);
            }
            break;
        // Add more cases for other text-based models as needed
        default:
            toastr.error('Unsupported model:', main_api);
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
            case 'koboldhorde':
                toastr.warning('Rewrite streaming not supported for Kobold. Turn off in rewrite settings.');
            default:
                throw new Error('Streaming is enabled, but the current API does not support streaming.');
        }
    } else {
        if (main_api === 'koboldhorde') {
            res = await generateHorde(prompt, generateData, abortController.signal, true);
        } else {
            const response = await generateRaw(prompt, null, false, false, null, generateData.max_length);
            res = {text: response};
            // Shamelessly copied from script.js
            /*function getGenerateUrl(api) {
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

            res = await response.json();*/
        }
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
        
        // Format the new text for proper display with paragraph spacing
        const messageData = getContext().chat[mesId];
        const formattedNewText = messageFormatting(
            newText,
            getContext().name2,
            messageData.isSystem,
            messageData.isUser,
            mesId
        );
        highlightedNewText.innerHTML = formattedNewText;

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
                case 'Instruct':
                    modifier = extension_settings[extensionName].instructTokensAdd;
                    break;
                case 'Custom1':
                    modifier = extension_settings[extensionName].custom1TokensAdd;
                    break;
                case 'Custom2':
                    modifier = extension_settings[extensionName].custom2TokensAdd;
                    break;
                case 'Custom3':
                    modifier = extension_settings[extensionName].custom3TokensAdd;
                    break;
                case 'Custom4':
                    modifier = extension_settings[extensionName].custom4TokensAdd;
                    break;
                case 'Custom5':
                    modifier = extension_settings[extensionName].custom5TokensAdd;
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
                case 'Instruct':
                    multiplier = extension_settings[extensionName].instructTokensMult;
                    break;
                case 'Custom1':
                    multiplier = extension_settings[extensionName].custom1TokensMult;
                    break;
                case 'Custom2':
                    multiplier = extension_settings[extensionName].custom2TokensMult;
                    break;
                case 'Custom3':
                    multiplier = extension_settings[extensionName].custom3TokensMult;
                    break;
                case 'Custom4':
                    multiplier = extension_settings[extensionName].custom4TokensMult;
                    break;
                case 'Custom5':
                    multiplier = extension_settings[extensionName].custom5TokensMult;
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
            case 'Instruct':
                result = extension_settings[extensionName].instructTokens;
                break;
            case 'Custom1':
                result = extension_settings[extensionName].custom1Tokens;
                break;
            case 'Custom2':
                result = extension_settings[extensionName].custom2Tokens;
                break;
            case 'Custom3':
                result = extension_settings[extensionName].custom3Tokens;
                break;
            case 'Custom4':
                result = extension_settings[extensionName].custom4Tokens;
                break;
            case 'Custom5':
                result = extension_settings[extensionName].custom5Tokens;
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

    // Apply AI Output regex scripts if setting is enabled
    let processedText = newText; // Default to original newText
    if (extension_settings[extensionName].applyRegexOnRewrite) {
        processedText = getRegexedString(newText, regex_placement.AI_OUTPUT);
    }

    // Create the new message with the rewritten and potentially processed section
    const newMessage =
        fullMessage.substring(0, startOffset) +
        processedText + // Use the processed text here
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
