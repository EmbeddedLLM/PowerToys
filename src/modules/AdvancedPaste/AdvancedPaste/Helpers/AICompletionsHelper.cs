// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using Azure;
using Azure.AI.OpenAI;
using ManagedCommon;
using Microsoft.PowerToys.Settings.UI.Library;
using Microsoft.PowerToys.Telemetry;
using Newtonsoft.Json.Linq;
using Windows.ApplicationModel.Chat;
using Windows.Security.Credentials;

namespace AdvancedPaste.Helpers
{
    public class AICompletionsHelper
    {
        // Return Response and Status code from the request.
        public struct AICompletionsResponse
        {
            public AICompletionsResponse(string response, int apiRequestStatus)
            {
                Response = response;
                ApiRequestStatus = apiRequestStatus;
            }

            public string Response { get; }

            public int ApiRequestStatus { get; }
        }

        private string _openAIKey;
        private string _localLLMEndpoint;

        private string _modelName = "gpt-3.5-turbo-instruct";

        public bool IsAIEnabled => !string.IsNullOrEmpty(this._openAIKey) || !string.IsNullOrEmpty(this._localLLMEndpoint);

        public AICompletionsHelper()
        {
            this._openAIKey = LoadOpenAIKey();
            this._localLLMEndpoint = LoadLocalLLMEndpoint();
        }

        public void SetOpenAIKey(string openAIKey)
        {
            this._openAIKey = openAIKey;
        }

        public void SetLocallLLMEndpoint(string localLLMEndpoint)
        {
            this._localLLMEndpoint = localLLMEndpoint;
        }

        public string GetLocallLLMEndpoint()
        {
            return _localLLMEndpoint;
        }

        public string GetKey()
        {
            return _openAIKey;
        }

        public static string LoadOpenAIKey()
        {
            PasswordVault vault = new PasswordVault();

            try
            {
                PasswordCredential cred = vault.Retrieve("https://platform.openai.com/api-keys", "PowerToys_AdvancedPaste_OpenAIKey");
                if (cred is not null)
                {
                    return cred.Password.ToString();
                }
            }
            catch (Exception)
            {
            }

            return string.Empty;
        }

        public static string LoadLocalLLMEndpoint()
        {
            PasswordVault vault = new PasswordVault();

            try
            {
                PasswordCredential cred = vault.Retrieve("localllm_endpoint", "PowerToys_AdvancedPaste_LocalLLMEndpoint");
                if (cred is not null)
                {
                    return cred.Password.ToString();
                }
            }
            catch (Exception)
            {
            }

            return string.Empty;
        }

        private Response<Completions> GetAICompletion(string systemInstructions, string userMessage)
        {
            OpenAIClient azureAIClient = new OpenAIClient(_openAIKey);
            int maxTokens = 2000;

            _modelName = "gpt-3.5-turbo-instruct";

            var completionsOptions = new CompletionsOptions()
            {
                DeploymentName = _modelName,
                Prompts = { systemInstructions + "\n\n" + userMessage },
                Temperature = 0.01F,
                MaxTokens = maxTokens,
            };
            var response = azureAIClient.GetCompletions(completionsOptions);

            if (response.Value.Choices[0].FinishReason == "length")
            {
                Console.WriteLine("Cut off due to length constraints");
            }

            return response;
        }

        private string GetAIChatCompletion(string systemInstructions, string userMessage)
        {
            int maxTokens = 256;

            // Create an instance of HttpClient
            using (HttpClient httpClient = new HttpClient())
            {
                // Set the timeout (e.g., 30 seconds)
                httpClient.Timeout = TimeSpan.FromSeconds(30);

                // Define the URL
                string url = $"{_localLLMEndpoint}/models";

                try
                {
                    // Perform the GET request asynchronously
                    HttpResponseMessage modelResponse = httpClient.GetAsync(url).Result;

                    // Ensure the request was successful
                    modelResponse.EnsureSuccessStatusCode();

                    // Read the response content
                    string responseBody = modelResponse.Content.ReadAsStringAsync().Result;

                    JObject jsonResponse = JObject.Parse(responseBody);
                    _modelName = jsonResponse["data"][0]["id"].ToString();

                    // Output the response
                    Console.WriteLine(responseBody);
                    Console.WriteLine(jsonResponse);
                }
                catch (HttpRequestException e)
                {
                    // Handle any errors that may have occurred
                    Console.WriteLine($"Request error: {e.Message}");
                    throw; // Re-throw the exception to be caught by the caller
                }
            }

            // Create an instance of HttpClient
            using (HttpClient httpClient = new HttpClient())
            {
                // Set the timeout (e.g., 30 seconds)
                httpClient.Timeout = TimeSpan.FromSeconds(30);

                try
                {
                    string chatUrl = $"{_localLLMEndpoint}/chat/completions";
                    var payload = new
                    {
                        messages = new[] { new { role = "system", content = systemInstructions }, new { role = "user", content = userMessage } },
                        model = _modelName,
                        max_tokens = maxTokens,
                        temperature = 0.0,
                        stream = false,
                    };
                    var jsonPayload = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
                    var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                    HttpResponseMessage chatResponse = httpClient.PostAsync(chatUrl, content).Result;

                    // Ensure the request was successful
                    chatResponse.EnsureSuccessStatusCode();
                    var responseText = chatResponse.Content.ReadAsStringAsync().Result;
                    Console.WriteLine(responseText);
                    JObject jsonResponseText = JObject.Parse(responseText);
                    return jsonResponseText["choices"][0]["message"]["content"].ToString();
                }
                catch (HttpRequestException e)
                {
                    // Handle any errors that may have occurred
                    Console.WriteLine($"Request error: {e.Message}");
                    throw; // Re-throw the exception to be caught by the caller
                }
            }
        }

        /* private Response<Completions> GetAICompletion(string systemInstructions, string userMessage)
        {
            OpenAIClient azureAIClient = null;
            int maxTokens = 256;

            if (!string.IsNullOrEmpty(_openAIKey))
            {
                azureAIClient = new OpenAIClient(_openAIKey);
                maxTokens = 2000;
                _modelName = "gpt-3.5-turbo-instruct";
            }
            else if (!string.IsNullOrEmpty(_localLLMEndpoint))
            {
                azureAIClient = new OpenAIClient(new Uri(_localLLMEndpoint), new AzureKeyCredential("EMPTY"));

                // Create an instance of HttpClient
                using (HttpClient client = new HttpClient())
                {
                    // Define the URL
                    string url = $"{_localLLMEndpoint}/models";

                    try
                    {
                        // Perform the GET request synchronously
                        HttpResponseMessage model_response = client.GetAsync(url).Result;

                        // Ensure the request was successful
                        model_response.EnsureSuccessStatusCode();

                        // Read the response content
                        string responseBody = model_response.Content.ReadAsStringAsync().Result;

                        JObject jsonResponse = JObject.Parse(responseBody);
                        _modelName = jsonResponse["data"][0]["id"].ToString();

                        // Output the response
                        Console.WriteLine(responseBody);

                        Console.WriteLine(jsonResponse);
                    }
                    catch (HttpRequestException e)
                    {
                        // Handle any errors that may have occurred
                        Console.WriteLine($"Request error: {e.Message}");
                    }
                }
            }

            var response = azureAIClient.GetCompletions(
                new CompletionsOptions()
                {
                    DeploymentName = _modelName,
                    Prompts =
                    {
                        systemInstructions + "\n\n" + userMessage,
                    },
                    Temperature = 0.01F,
                    MaxTokens = maxTokens,
                });

            if (response.Value.Choices[0].FinishReason == "length")
            {
                Console.WriteLine("Cut off due to length constraints");
            }

            return response;
        }*/

        public AICompletionsResponse AIFormatString(string inputInstructions, string inputString)
        {
            string systemInstructions = $@"You are tasked with user's clipboard data. Use the user's instructions, and the content of their clipboard below to edit their clipboard content as they have requested it.

Do not output anything else besides the reformatted clipboard content.";

            string userMessage = $@"User instructions:
{inputInstructions}

Clipboard Content:
{inputString}

Output:
";

            string aiResponse = null;
            Response<Completions> rawAIResponse = null;

            // Response<ChatCompletions> rawAIChatResponse = null;
            int apiRequestStatus = (int)HttpStatusCode.OK;

            try
            {
                if (!string.IsNullOrEmpty(_openAIKey))
                {
                    rawAIResponse = this.GetAICompletion(systemInstructions, userMessage);
                    aiResponse = rawAIResponse.Value.Choices[0].Text;

                    int promptTokens = rawAIResponse.Value.Usage.PromptTokens;
                    int completionTokens = rawAIResponse.Value.Usage.CompletionTokens;
                    PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomFormatEvent(promptTokens, completionTokens, _modelName));
                    return new AICompletionsResponse(aiResponse, apiRequestStatus);
                }
                else if (!string.IsNullOrEmpty(_localLLMEndpoint))
                {
                    aiResponse = this.GetAIChatCompletion(systemInstructions, userMessage);

                    int promptTokens = -1; // rawAIChatResponse.Value.Usage.PromptTokens;
                    int completionTokens = -1; // rawAIChatResponse.Value.Usage.CompletionTokens;
                    PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomFormatEvent(promptTokens, completionTokens, _modelName));
                    return new AICompletionsResponse(aiResponse, apiRequestStatus);
                }
            }
            catch (Azure.RequestFailedException error)
            {
                Logger.LogError("GetAICompletion failed", error);
                PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomErrorEvent(error.Message));
                apiRequestStatus = error.Status;
            }
            catch (Exception error)
            {
                Logger.LogError("GetAICompletion failed", error);
                PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomErrorEvent(error.Message));
                apiRequestStatus = -1;
            }

            return new AICompletionsResponse(aiResponse, apiRequestStatus);
        }

        public AICompletionsResponse AISIIFormatString(string systemInstructions, string inputInstructions, string inputString)
        {
            string userMessage = $"{inputInstructions}\n{inputString}";

            string aiResponse = null;
            Response<Completions> rawAIResponse = null;

            // Response<ChatCompletions> rawAIChatResponse = null;
            int apiRequestStatus = (int)HttpStatusCode.OK;

            try
            {
                if (!string.IsNullOrEmpty(_openAIKey))
                {
                    rawAIResponse = this.GetAICompletion(systemInstructions, userMessage);
                    aiResponse = rawAIResponse.Value.Choices[0].Text;

                    int promptTokens = rawAIResponse.Value.Usage.PromptTokens;
                    int completionTokens = rawAIResponse.Value.Usage.CompletionTokens;
                    PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomFormatEvent(promptTokens, completionTokens, _modelName));
                    return new AICompletionsResponse(aiResponse, apiRequestStatus);
                }
                else if (!string.IsNullOrEmpty(_localLLMEndpoint))
                {
                    aiResponse = this.GetAIChatCompletion(systemInstructions, userMessage);

                    int promptTokens = -1; // rawAIChatResponse.Value.Usage.PromptTokens;
                    int completionTokens = -1; // rawAIChatResponse.Value.Usage.CompletionTokens;
                    PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomFormatEvent(promptTokens, completionTokens, _modelName));
                    return new AICompletionsResponse(aiResponse, apiRequestStatus);
                }
            }
            catch (Azure.RequestFailedException error)
            {
                Logger.LogError("GetAICompletion failed", error);
                PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomErrorEvent(error.Message));
                apiRequestStatus = error.Status;
            }
            catch (Exception error)
            {
                Logger.LogError("GetAICompletion failed", error);
                PowerToysTelemetry.Log.WriteEvent(new Telemetry.AdvancedPasteGenerateCustomErrorEvent(error.Message));
                apiRequestStatus = -1;
            }

            return new AICompletionsResponse(aiResponse, apiRequestStatus);
        }
    }
}
