// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Text.Json;
using System.Text.Json.Serialization;
using Settings.UI.Library.Attributes;

namespace Microsoft.PowerToys.Settings.UI.Library
{
    public class AdvancedPasteProperties
    {
        public static readonly HotkeySettings DefaultAdvancedPasteUIShortcut = new HotkeySettings(true, false, false, true, 0x56); // Win+Shift+V

        public static readonly HotkeySettings DefaultPasteAsPlainTextShortcut = new HotkeySettings(true, true, true, false, 0x56); // Ctrl+Win+Alt+V

        // public static readonly string DefaultLocalLLMEndpoint = "http://localhost:6979/v1";
        public AdvancedPasteProperties()
        {
            AdvancedPasteUIShortcut = DefaultAdvancedPasteUIShortcut;
            PasteAsPlainTextShortcut = DefaultPasteAsPlainTextShortcut;
            PasteAsMarkdownShortcut = new();
            PasteAsJsonShortcut = new();

            // CHANGED
            RewriteProfessionallyAndPasteAsPlainTextShortcut = new();
            ShowCustomPreview = true;
            SendPasteKeyCombination = true;
            CloseAfterLosingFocus = false;
            LocalLLMEndpoint = string.Empty;
        }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool ShowCustomPreview { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        [CmdConfigureIgnore]
        public bool SendPasteKeyCombination { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool CloseAfterLosingFocus { get; set; }

        [JsonPropertyName("local-llm-endpoint")]
        public string LocalLLMEndpoint { get; set; }

        [JsonPropertyName("advanced-paste-ui-hotkey")]
        public HotkeySettings AdvancedPasteUIShortcut { get; set; }

        [JsonPropertyName("paste-as-plain-hotkey")]
        public HotkeySettings PasteAsPlainTextShortcut { get; set; }

        [JsonPropertyName("paste-as-markdown-hotkey")]
        public HotkeySettings PasteAsMarkdownShortcut { get; set; }

        [JsonPropertyName("paste-as-json-hotkey")]
        public HotkeySettings PasteAsJsonShortcut { get; set; }

        // CHANGED
        [JsonPropertyName("rewrite-professionaly-and-paste-as-plain-hotkey")]
        public HotkeySettings RewriteProfessionallyAndPasteAsPlainTextShortcut { get; set; }

        public override string ToString()
            => JsonSerializer.Serialize(this);
    }
}
