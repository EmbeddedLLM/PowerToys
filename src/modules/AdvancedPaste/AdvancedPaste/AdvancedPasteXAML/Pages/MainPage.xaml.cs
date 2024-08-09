// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AdvancedPaste.Helpers;
using AdvancedPaste.Models;
using AdvancedPaste.ViewModels;
using ManagedCommon;
using Microsoft.PowerToys.Telemetry;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Streams;
using Windows.System;

namespace AdvancedPaste.Pages
{
    public sealed partial class MainPage : Page
    {
        private readonly ObservableCollection<ClipboardItem> clipboardHistory;
        private readonly ObservableCollection<PasteFormat> pasteFormats;
        private readonly Microsoft.UI.Dispatching.DispatcherQueue _dispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();

        public OptionsViewModel ViewModel { get; private set; }

        public MainPage()
        {
            this.InitializeComponent();

            pasteFormats =
            [
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE8E9" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("PasteAsPlainText"), Format = PasteFormats.PlainText },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\ue8a5" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("PasteAsMarkdown"), Format = PasteFormats.Markdown },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE943" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("PasteAsJson"), Format = PasteFormats.Json },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("ProofreadAndPasteAsPlainText"), Format = PasteFormats.Proofread },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("RewriteAndPasteAsPlainText"), Format = PasteFormats.Rewrite },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("RewriteProfessionallyAndPasteAsPlainText"), Format = PasteFormats.WriteProfessionally },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("RewriteConciselyAndPasteAsPlainText"), Format = PasteFormats.WriteConcisely },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("RewriteFriendlyAndPasteAsPlainText"), Format = PasteFormats.WriteFriendly },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("SummarizeAndPasteAsPlainText"), Format = PasteFormats.Summarize },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("SummarizeToKeyPointsAndPasteAsPlainText"), Format = PasteFormats.SummarizeToKeyPoints },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("SummarizeToTableAndPasteAsPlainText"), Format = PasteFormats.SummarizeToTable },
                new PasteFormat { Icon = new FontIcon() { Glyph = "\uE945" }, Name = ResourceLoaderInstance.ResourceLoader.GetString("SummarizeToListAndPasteAsPlainText"), Format = PasteFormats.SummarizeToList },
            ];

            ViewModel = App.GetService<OptionsViewModel>();

            clipboardHistory = new ObservableCollection<ClipboardItem>();

            LoadClipboardHistoryEvent(null, null);
            Clipboard.HistoryChanged += LoadClipboardHistoryEvent;
        }

        private void LoadClipboardHistoryEvent(object sender, object e)
        {
            Task.Run(() =>
            {
                LoadClipboardHistoryAsync();
            });
        }

        public async void LoadClipboardHistoryAsync()
        {
            try
            {
                Logger.LogTrace();

                List<ClipboardItem> items = new();

                if (Clipboard.IsHistoryEnabled())
                {
                    var historyItems = await Clipboard.GetHistoryItemsAsync();
                    if (historyItems.Status == ClipboardHistoryItemsResultStatus.Success)
                    {
                        foreach (var item in historyItems.Items)
                        {
                            if (item.Content.Contains(StandardDataFormats.Text))
                            {
                                string text = await item.Content.GetTextAsync();
                                items.Add(new ClipboardItem { Content = text, Item = item });
                            }
                            else if (item.Content.Contains(StandardDataFormats.Bitmap))
                            {
                                items.Add(new ClipboardItem { Item = item });
                            }
                        }
                    }
                }

                _dispatcherQueue.TryEnqueue(async () =>
                {
                    // Clear to avoid leaks due to Garbage Collection not clearing the bitmap from memory. Fix for https://github.com/microsoft/PowerToys/issues/33423
                    clipboardHistory.Where(x => x.Image is not null)
                                    .ToList()
                                    .ForEach(x => x.Image.ClearValue(BitmapImage.UriSourceProperty));

                    clipboardHistory.Clear();

                    foreach (var item in items)
                    {
                        if (item.Item.Content.Contains(StandardDataFormats.Bitmap))
                        {
                            IRandomAccessStreamReference imageReceived = null;
                            imageReceived = await item.Item.Content.GetBitmapAsync();
                            if (imageReceived != null)
                            {
                                using (var imageStream = await imageReceived.OpenReadAsync())
                                {
                                    var bitmapImage = new BitmapImage();
                                    bitmapImage.SetSource(imageStream);
                                    item.Image = bitmapImage;
                                }
                            }
                        }

                        clipboardHistory.Add(item);
                    }
                });
            }
            catch (Exception ex)
            {
                Logger.LogError("Loading clipboard history failed", ex);
            }
        }

        private void ClipboardHistoryItemDeleteButton_Click(object sender, RoutedEventArgs e)
        {
            Logger.LogTrace();

            if (sender is MenuFlyoutItem btn)
            {
                ClipboardItem item = btn.CommandParameter as ClipboardItem;
                Clipboard.DeleteItemFromHistory(item.Item);
                clipboardHistory.Remove(item);

                PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteClipboardItemDeletedEvent());
            }
        }

        private void PasteAsPlain()
        {
            ViewModel.ToPlainTextFunction();
        }

        private void PasteAsMarkdown()
        {
            ViewModel.ToMarkdownFunction();
        }

        private void PasteAsJson()
        {
            ViewModel.ToJsonFunction();
        }

        private void Proofread()
        {
            ViewModel.ProofreadFunction();
        }

        private void Rewrite()
        {
            ViewModel.RewriteFunction();
        }

        private void RewriteProfessionally()
        {
            ViewModel.RewriteProfessionallyFunction();
        }

        private void RewriteConcisely()
        {
            ViewModel.RewriteConciselyFunction();
        }

        private void RewriteFriendly()
        {
            ViewModel.RewriteFriendlyFunction();
        }

        private void Summarize()
        {
            ViewModel.SummarizeFunction();
        }

        private void SummarizeToKeyPoints()
        {
            ViewModel.SummarizeToKeyPointsFunction();
        }

        private void SummarizeToTable()
        {
            ViewModel.SummarizeToTableFunction();
        }

        private void SummarizeToList()
        {
            ViewModel.SummarizeToListFunction();
        }

        private void PasteOptionsListView_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (e.ClickedItem is PasteFormat format)
            {
                switch (format.Format)
                {
                    case PasteFormats.PlainText:
                        {
                            PasteAsPlain();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.PlainText));
                            break;
                        }

                    case PasteFormats.Markdown:
                        {
                            PasteAsMarkdown();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.Markdown));
                            break;
                        }

                    case PasteFormats.Json:
                        {
                            PasteAsJson();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.Json));
                            break;
                        }

                    case PasteFormats.Rewrite:
                        {
                            // Placeholder
                            Rewrite();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.Rewrite));
                            break;
                        }

                    case PasteFormats.Proofread:
                        {
                            // Placeholder
                            Proofread();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.Proofread));
                            break;
                        }

                    case PasteFormats.WriteProfessionally:
                        {
                            // Placeholder
                            RewriteProfessionally();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                            break;
                        }

                    case PasteFormats.WriteConcisely:
                        {
                            // Placeholder
                            RewriteConcisely();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteConcisely));
                            break;
                        }

                    case PasteFormats.WriteFriendly:
                        {
                            // Placeholder
                            RewriteFriendly();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteFriendly));
                            break;
                        }

                    case PasteFormats.Summarize:
                        {
                            // Placeholder
                            Summarize();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.Summarize));
                            break;
                        }

                    case PasteFormats.SummarizeToKeyPoints:
                        {
                            // Placeholder
                            SummarizeToKeyPoints();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.SummarizeToKeyPoints));
                            break;
                        }

                    case PasteFormats.SummarizeToTable:
                        {
                            // Placeholder
                            SummarizeToTable();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.SummarizeToTable));
                            break;
                        }

                    case PasteFormats.SummarizeToList:
                        {
                            // Placeholder
                            SummarizeToList();
                            PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.SummarizeToList));
                            break;
                        }
                }
            }
        }

        private void KeyboardAccelerator_Invoked(Microsoft.UI.Xaml.Input.KeyboardAccelerator sender, Microsoft.UI.Xaml.Input.KeyboardAcceleratorInvokedEventArgs args)
        {
            Logger.LogTrace();

            switch (sender.Key)
            {
                case VirtualKey.Escape:
                    {
                        (App.Current as App).GetMainWindow().Close();
                        break;
                    }

                case VirtualKey.Number1:
                    {
                        PasteAsPlain();
                        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteInAppKeyboardShortcutEvent(PasteFormats.PlainText));
                        break;
                    }

                case VirtualKey.Number2:
                    {
                        PasteAsMarkdown();
                        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteInAppKeyboardShortcutEvent(PasteFormats.Markdown));
                        break;
                    }

                case VirtualKey.Number3:
                    {
                        PasteAsJson();
                        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteInAppKeyboardShortcutEvent(PasteFormats.Json));
                        break;
                    }

                // case VirtualKey.Number4:
                //    {
                //        // Placeholder
                //        Proofread();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }

                // case VirtualKey.Number5:
                //    {
                //        // Placeholder
                //        RewriteProfessionally();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }

                // case VirtualKey.Number6:
                //    {
                //        // Placeholder
                //        RewriteConcisely();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }

                // case VirtualKey.Number7:
                //    {
                //        // Placeholder
                //        RewriteFriendly();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }

                // case VirtualKey.Number8:
                //    {
                //        // Placeholder
                //        Summarize();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }

                // case VirtualKey.Number9:
                //    {
                //        // Placeholder
                //        SummarizeToKeyPoints();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }

                // case VirtualKey.Number10:
                //    {
                //        // Placeholder
                //        SummarizeToKeyPoints();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }

                // case VirtualKey.Number11:
                //    {
                //        // Placeholder
                //        SummarizeToKeyPoints();
                //        PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteFormatClickedEvent(PasteFormats.WriteProfessionally));
                //        break;
                //    }
                default:
                    break;
            }
        }

        private void Page_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
        {
            if (e.Key == VirtualKey.Escape)
            {
                (App.Current as App).GetMainWindow().Close();
            }
        }

        private async void ClipboardHistory_ItemClick(object sender, ItemClickEventArgs e)
        {
            var item = e.ClickedItem as ClipboardItem;
            if (item is not null)
            {
                PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteClipboardItemClicked());
                if (!string.IsNullOrEmpty(item.Content))
                {
                    ClipboardHelper.SetClipboardTextContent(item.Content);
                }
                else if (item.Image is not null)
                {
                    RandomAccessStreamReference image = null;
                    image = await item.Item.Content.GetBitmapAsync();
                    ClipboardHelper.SetClipboardImageContent(image);
                }
            }
        }
    }
}
