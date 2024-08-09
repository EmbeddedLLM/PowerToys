// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Net;
using System.Threading.Tasks;
using AdvancedPaste.Helpers;
using AdvancedPaste.Models;
using AdvancedPaste.Settings;
using Common.UI;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ManagedCommon;
using Microsoft.PowerToys.Settings.UI.Library;
using Microsoft.UI.Dispatching;
using Microsoft.Win32;
using Windows.ApplicationModel.DataTransfer;
using WinUIEx;

namespace AdvancedPaste.ViewModels
{
    public partial class OptionsViewModel : ObservableObject
    {
        private readonly DispatcherQueue _dispatcherQueue = DispatcherQueue.GetForCurrentThread();
        private readonly IUserSettings _userSettings;

        private App app = App.Current as App;

        private AICompletionsHelper aiHelper;

        public DataPackageView ClipboardData { get; set; }

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(InputTxtBoxPlaceholderText))]
        [NotifyPropertyChangedFor(nameof(IsCustomAIEnabled))]
        private bool _isClipboardDataText;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(InputTxtBoxPlaceholderText))]
        private bool _isCustomAIEnabled;

        [ObservableProperty]
        private bool _clipboardHistoryEnabled;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(InputTxtBoxErrorText))]
        private int _apiRequestStatus;

        public OptionsViewModel(IUserSettings userSettings)
        {
            aiHelper = new AICompletionsHelper();
            _userSettings = userSettings;

            IsCustomAIEnabled = IsClipboardDataText && aiHelper.IsAIEnabled;

            ApiRequestStatus = (int)HttpStatusCode.OK;

            GeneratedResponses = new ObservableCollection<string>();
            GeneratedResponses.CollectionChanged += (s, e) =>
            {
                OnPropertyChanged(nameof(HasMultipleResponses));
                OnPropertyChanged(nameof(CurrentIndexDisplay));
            };

            ClipboardHistoryEnabled = IsClipboardHistoryEnabled();
            GetClipboardData();
        }

        public void GetClipboardData()
        {
            ClipboardData = Clipboard.GetContent();
            IsClipboardDataText = ClipboardData.Contains(StandardDataFormats.Text);
        }

        public void OnShow()
        {
            GetClipboardData();

            if (PowerToys.GPOWrapper.GPOWrapper.GetAllowedAdvancedPasteOnlineAIModelsValue() == PowerToys.GPOWrapper.GpoRuleConfigured.Disabled)
            {
                IsCustomAIEnabled = false;
                OnPropertyChanged(nameof(InputTxtBoxPlaceholderText));
            }
            else
            {
                var openAIKey = AICompletionsHelper.LoadOpenAIKey();
                var currentKey = aiHelper.GetKey();
                var localLLMEndpoint = AICompletionsHelper.LoadLocalLLMEndpoint();
                var currentLocalLLMEndpoint = aiHelper.GetLocallLLMEndpoint();
                bool keyChanged = (openAIKey != currentKey) || (localLLMEndpoint != currentLocalLLMEndpoint);

                if (keyChanged)
                {
                    app.GetMainWindow().StartLoading();

                    Task.Run(() =>
                    {
                        aiHelper.SetOpenAIKey(openAIKey);
                        aiHelper.SetLocallLLMEndpoint(localLLMEndpoint);
                    }).ContinueWith(
                        (t) =>
                        {
                            _dispatcherQueue.TryEnqueue(() =>
                            {
                                app.GetMainWindow().FinishLoading(aiHelper.IsAIEnabled);
                                OnPropertyChanged(nameof(InputTxtBoxPlaceholderText));
                                IsCustomAIEnabled = IsClipboardDataText && aiHelper.IsAIEnabled;
                            });
                        },
                        TaskScheduler.Default);
                }
                else
                {
                    IsCustomAIEnabled = IsClipboardDataText && aiHelper.IsAIEnabled;
                }
            }

            ClipboardHistoryEnabled = IsClipboardHistoryEnabled();
            GeneratedResponses.Clear();
        }

        // List to store generated responses
        public ObservableCollection<string> GeneratedResponses { get; set; } = new ObservableCollection<string>();

        // Index to keep track of the current response
        private int _currentResponseIndex;

        public int CurrentResponseIndex
        {
            get => _currentResponseIndex;
            set
            {
                if (value >= 0 && value < GeneratedResponses.Count)
                {
                    SetProperty(ref _currentResponseIndex, value);
                    CustomFormatResult = GeneratedResponses[_currentResponseIndex];
                    OnPropertyChanged(nameof(CurrentIndexDisplay));
                }
            }
        }

        public bool HasMultipleResponses
        {
            get => GeneratedResponses.Count > 1;
        }

        public string CurrentIndexDisplay => $"{CurrentResponseIndex + 1}/{GeneratedResponses.Count}";

        public string InputTxtBoxPlaceholderText
        {
            get
            {
                app.GetMainWindow().ClearInputText();

                if (PowerToys.GPOWrapper.GPOWrapper.GetAllowedAdvancedPasteOnlineAIModelsValue() == PowerToys.GPOWrapper.GpoRuleConfigured.Disabled)
                {
                    return ResourceLoaderInstance.ResourceLoader.GetString("OpenAIGpoDisabled");
                }
                else if (!aiHelper.IsAIEnabled)
                {
                    return ResourceLoaderInstance.ResourceLoader.GetString("OpenAINotConfigured");
                }
                else if (!IsClipboardDataText)
                {
                    return ResourceLoaderInstance.ResourceLoader.GetString("ClipboardDataTypeMismatchWarning");
                }
                else
                {
                    return ResourceLoaderInstance.ResourceLoader.GetString("CustomFormatTextBox/PlaceholderText");
                }
            }
        }

        public string InputTxtBoxErrorText
        {
            get
            {
                if (ApiRequestStatus != (int)HttpStatusCode.OK)
                {
                    if (ApiRequestStatus == (int)HttpStatusCode.TooManyRequests)
                    {
                        return ResourceLoaderInstance.ResourceLoader.GetString("OpenAIApiKeyTooManyRequests");
                    }
                    else if (ApiRequestStatus == (int)HttpStatusCode.Unauthorized)
                    {
                        return ResourceLoaderInstance.ResourceLoader.GetString("OpenAIApiKeyUnauthorized");
                    }
                    else
                    {
                        return ResourceLoaderInstance.ResourceLoader.GetString("OpenAIApiKeyError") + ApiRequestStatus.ToString(CultureInfo.InvariantCulture);
                    }
                }

                return string.Empty;
            }
        }

        [ObservableProperty]
        private string _customFormatResult;

        [RelayCommand]
        public void PasteCustom()
        {
            PasteCustomFunction(GeneratedResponses[CurrentResponseIndex]);
        }

        // Command to select the previous custom format
        [RelayCommand]
        public void PreviousCustomFormat()
        {
            if (CurrentResponseIndex > 0)
            {
                CurrentResponseIndex--;
            }
        }

        // Command to select the next custom format
        [RelayCommand]
        public void NextCustomFormat()
        {
            if (CurrentResponseIndex < GeneratedResponses.Count - 1)
            {
                CurrentResponseIndex++;
            }
        }

        // Command to open the Settings window.
        [RelayCommand]
        public void OpenSettings()
        {
            SettingsDeepLink.OpenSettings(SettingsDeepLink.SettingsWindow.AdvancedPaste, true);
            (App.Current as App).GetMainWindow().Close();
        }

        private void SetClipboardContentAndHideWindow(string content)
        {
            if (!string.IsNullOrEmpty(content))
            {
                ClipboardHelper.SetClipboardTextContent(content);
            }

            if (app.GetMainWindow() != null)
            {
                Windows.Win32.Foundation.HWND hwnd = (Windows.Win32.Foundation.HWND)app.GetMainWindow().GetWindowHandle();
                Windows.Win32.PInvoke.ShowWindow(hwnd, Windows.Win32.UI.WindowsAndMessaging.SHOW_WINDOW_CMD.SW_HIDE);
            }
        }

        internal void ToPlainTextFunction()
        {
            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void ToMarkdownFunction(bool pasteAlways = false)
        {
            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.ToMarkdown(ClipboardData);

                SetClipboardContentAndHideWindow(outputString);

                if (pasteAlways || _userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void ToJsonFunction(bool pasteAlways = false)
        {
            try
            {
                Logger.LogTrace();

                string jsonText = JsonHelper.ToJsonFromXmlOrCsv(ClipboardData);

                SetClipboardContentAndHideWindow(jsonText);

                if (pasteAlways || _userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void ProofreadFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will carefully review the text for spelling, grammar, and punctuation errors, then provide corrections and suggestions for improvement.";
                    string inputInstructions = "Please proofread the following text for errors:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void RewriteFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will rephrase the given text to enhance its readability and coherence while maintaining the original meaning.";
                    string inputInstructions = "Rewrite this text to improve its clarity and flow:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void RewriteProfessionallyFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will refine the language to be more formal, precise, and suitable for a professional context.";
                    string inputInstructions = "Adjust this text to have a more professional tone:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void RewriteConciselyFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will condense the given text, removing unnecessary details while preserving the core message and key information.";
                    string inputInstructions = "Shorten this text while keeping the main points:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void RewriteFriendlyFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will revise the text to adopt a warmer, more conversational tone that engages the reader in a friendly manner.";
                    string inputInstructions = "Make this text sound more friendly and approachable:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void SummarizeFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will create a concise overview that captures the main ideas and essential points of the given text.";
                    string inputInstructions = "Provide a brief summary of the following text:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void SummarizeToKeyPointsFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will identify and list the most important ideas and information from the given text in a clear, bullet-point format.";
                    string inputInstructions = "Extract the key points from this text:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void SummarizeToTableFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will structure the given information into a clear, easy-to-read table with appropriate columns and rows.";
                    string inputInstructions = "Organize the information in this text into a table format:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal void SummarizeToListFunction(bool pasteAlways = false)
        {
            Logger.LogTrace();

            try
            {
                Logger.LogTrace();

                string outputString = MarkdownHelper.PasteAsPlainTextFromClipboard(ClipboardData);

                if (!string.IsNullOrWhiteSpace(outputString))
                {
                    Logger.LogWarning("Clipboard has no usable text data");

                    string systemInstructions = "I will reorganize the information from the given text into a bulleted or numbered list format for easier reading and comprehension.";
                    string inputInstructions = "Convert this text into a structured list:";
                    var aiResponse = aiHelper.AISIIFormatString(systemInstructions, inputInstructions, outputString);

                    outputString = aiResponse.Response;
                }

                SetClipboardContentAndHideWindow(outputString);

                if (_userSettings.SendPasteKeyCombination)
                {
                    ClipboardHelper.SendPasteKeyCombination();
                }
            }
            catch
            {
            }
        }

        internal async Task<string> GenerateCustomFunction(string inputInstructions)
        {
            Logger.LogTrace();

            if (string.IsNullOrWhiteSpace(inputInstructions))
            {
                return string.Empty;
            }

            if (ClipboardData == null || !ClipboardData.Contains(StandardDataFormats.Text))
            {
                Logger.LogWarning("Clipboard does not contain text data");
                return string.Empty;
            }

            string currentClipboardText = await Task.Run(async () =>
            {
                try
                {
                    string text = await ClipboardData.GetTextAsync() as string;
                    return text;
                }
                catch (Exception)
                {
                    // Couldn't get text from the clipboard. Resume with empty text.
                    return string.Empty;
                }
            });

            if (string.IsNullOrWhiteSpace(currentClipboardText))
            {
                Logger.LogWarning("Clipboard has no usable text data");
                return string.Empty;
            }

            var aiResponse = await Task.Run(() => aiHelper.AIFormatString(inputInstructions, currentClipboardText));

            string aiOutput = aiResponse.Response;
            ApiRequestStatus = aiResponse.ApiRequestStatus;

            GeneratedResponses.Add(aiOutput);
            CurrentResponseIndex = GeneratedResponses.Count - 1;
            return aiOutput;
        }

        internal void PasteCustomFunction(string text)
        {
            Logger.LogTrace();

            SetClipboardContentAndHideWindow(text);

            if (_userSettings.SendPasteKeyCombination)
            {
                ClipboardHelper.SendPasteKeyCombination();
            }
        }

        internal CustomQuery RecallPreviousCustomQuery()
        {
            return LoadPreviousQuery();
        }

        internal void SaveQuery(string inputQuery)
        {
            Logger.LogTrace();

            DataPackageView clipboardData = Clipboard.GetContent();

            if (clipboardData == null || !clipboardData.Contains(StandardDataFormats.Text))
            {
                Logger.LogWarning("Clipboard does not contain text data");
                return;
            }

            string currentClipboardText = Task.Run(async () =>
            {
                string text = await clipboardData.GetTextAsync() as string;
                return text;
            }).Result;

            var queryData = new CustomQuery
            {
                Query = inputQuery,
                ClipboardData = currentClipboardText,
            };

            SettingsUtils utils = new SettingsUtils();
            utils.SaveSettings(queryData.ToString(), Constants.AdvancedPasteModuleName, Constants.LastQueryJsonFileName);
        }

        internal CustomQuery LoadPreviousQuery()
        {
            SettingsUtils utils = new SettingsUtils();
            var query = utils.GetSettings<CustomQuery>(Constants.AdvancedPasteModuleName, Constants.LastQueryJsonFileName);
            return query;
        }

        private bool IsClipboardHistoryEnabled()
        {
            string registryKey = @"HKEY_CURRENT_USER\Software\Microsoft\Clipboard\";
            try
            {
                int enableClipboardHistory = (int)Registry.GetValue(registryKey, "EnableClipboardHistory", false);
                return enableClipboardHistory != 0;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
