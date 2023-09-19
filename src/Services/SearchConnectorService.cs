// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Azure.Identity;
using Microsoft.Graph;

// Using beta SDK to get access to the iconUrl value in the labels enum
using Microsoft.Graph.Beta;
using Microsoft.Graph.Beta.External.Connections.Item.Items.Item.MicrosoftGraphExternalConnectorsAddActivities;
using Microsoft.Graph.Beta.Models.ExternalConnectors;
using Microsoft.Kiota.Authentication.Azure;

namespace GitHubConnector.Services;

/// <summary>
/// Contains methods to manage search connections, schema, and items.
/// </summary>
public class SearchConnectorService
{
    /// <summary>
    /// Schema for ingesting GitHub issues.
    /// </summary>
    public static readonly Schema IssuesSchema = new()
    {
        BaseType = "microsoft.graph.externalItem",
        Properties = new()
        {
            new() { Aliases = new() { "issueTitle" }, Name = "title", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true,  IsRefinable = false, Labels = new() { Label.Title } },
            new() { Aliases = new() { "message" }, Name = "body", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false },
            new() { Name = "assignees", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false },
            new() { Name = "labels", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false },
            new() { Name = "state", Type = PropertyType.String, IsSearchable = false, IsQueryable = true, IsRetrievable = true, IsRefinable = true },
            new() { Name = "issueUrl", Type = PropertyType.String, IsSearchable = false, IsQueryable = false, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.Url } },
            new() { Name = "icon", Type = PropertyType.String, IsSearchable = false, IsQueryable = false, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.IconUrl } },
            new() { Name = "updatedAt", Type = PropertyType.DateTime, IsSearchable = false, IsQueryable = true, IsRetrievable = true, IsRefinable = true, Labels = new() { Label.LastModifiedDateTime } },
            new() { Name = "lastModifiedBy", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.LastModifiedBy } },
        },
    };

    /// <summary>
    /// Schema for ingesting GitHub repositories.
    /// </summary>
    public static readonly Schema ReposSchema = new()
    {
        BaseType = "microsoft.graph.externalItem",
        Properties = new()
        {
            new() { Aliases = new() { "repoName" }, Name = "title", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.Title } },
            new() { Name = "description", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false },
            new() { Name = "visibility", Type = PropertyType.String, IsSearchable = false, IsQueryable = true, IsRetrievable = true, IsRefinable = true },
            new() { Name = "createdBy", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.CreatedBy } },
            new() { Name = "updatedAt", Type = PropertyType.DateTime, IsSearchable = false, IsQueryable = true, IsRetrievable = true, IsRefinable = true, Labels = new() { Label.LastModifiedDateTime } },
            new() { Name = "lastModifiedBy", Type = PropertyType.String, IsSearchable = true, IsQueryable = true, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.LastModifiedBy } },
            new() { Name = "repoUrl", Type = PropertyType.String, IsSearchable = true, IsQueryable = false, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.Url } },
            new() { Name = "userUrl", Type = PropertyType.String, IsSearchable = true, IsQueryable = false, IsRetrievable = true, IsRefinable = false },
            new() { Name = "icon", Type = PropertyType.String, IsSearchable = false, IsQueryable = false, IsRetrievable = true, IsRefinable = false, Labels = new() { Label.IconUrl } },
        },
    };

    private readonly GraphServiceClient graphClient;
    private readonly HttpClient httpClient;
    private readonly string gitHubOwner;
    private readonly string gitHubRepo;
    private readonly string placeholderUserId;

    /// <summary>
    /// Initializes a new instance of the <see cref="SearchConnectorService"/> class.
    /// </summary>
    /// <param name="settings">The application settings.</param>
    /// <exception cref="ArgumentException">Thrown if required settings are not contained in the settings argument.</exception>
    public SearchConnectorService(AppSettings settings)
    {
        gitHubOwner = !string.IsNullOrEmpty(settings.GitHubRepoOwner) ?
            settings.GitHubRepoOwner : throw new ArgumentException("gitHubRepoOwner not set in app settings");
        gitHubRepo = !string.IsNullOrEmpty(settings.GitHubRepo) ?
            settings.GitHubRepo : throw new ArgumentException("gitHubRepo not set in app settings");
        placeholderUserId = !string.IsNullOrEmpty(settings.PlaceholderUserId) ?
            settings.PlaceholderUserId : throw new ArgumentException("placeholderUserId not set in app settings");

        httpClient = GraphClientFactory.Create();

        var credential = new ClientSecretCredential(
            settings.TenantId, settings.ClientId, settings.ClientSecret);

        var authProvider = new AzureIdentityAuthenticationProvider(
            credential, scopes: new[] { "https://graph.microsoft.com/.default" });

        graphClient = new GraphServiceClient(httpClient, authProvider);
    }

    /// <summary>
    /// Creates an <see cref="ExternalConnection"/>.
    /// </summary>
    /// <param name="connectionId">The connection ID for the new connection.</param>
    /// <param name="name">The display name of the new connection.</param>
    /// <param name="description">The description of the new connection.</param>
    /// <param name="itemType">The item type for the new connection (`issues` or `repos`).</param>
    /// <returns>The new <see cref="ExternalConnection"/>.</returns>
    public Task<ExternalConnection?> CreateConnectionAsync(string connectionId, string name, string? description, string itemType)
    {
        var newConnection = new ExternalConnection
        {
            Id = connectionId,
            Name = name,
            Description = description,
            ActivitySettings = new()
            {
                UrlToItemResolvers = new()
                {
                    new ItemIdResolver
                    {
                        Priority = 1,
                        ItemId = itemType == "issues" ? "{issueId}" : "{repo}",
                        UrlMatchInfo = new()
                        {
                            UrlPattern = itemType == "issues" ?
                                $"/{gitHubOwner}/{gitHubRepo}/issues/(?<issueId>[0-9]+)" :
                                $"/{gitHubOwner}/(?<repo>.*)/",
                            BaseUrls = new() { "https://github.com" },
                        },
                    },
                },
            },
        };

        return graphClient.External.Connections.PostAsync(newConnection);
    }

    /// <summary>
    /// Gets existing connections.
    /// </summary>
    /// <returns>An <see cref="ExternalConnectionCollectionResponse"/> containing the existing connections.</returns>
    public Task<ExternalConnectionCollectionResponse?> GetConnectionsAsync()
    {
        return graphClient.External.Connections.GetAsync();
    }

    /// <summary>
    /// Delete a connection.
    /// </summary>
    /// <param name="connectionId">The connection ID of the connection to delete.</param>
    /// <returns>A <see cref="Task"/> indicating the status of the asynchronous delete operation.</returns>
    public Task DeleteConnectionAsync(string connectionId)
    {
        return graphClient.External.Connections[connectionId].DeleteAsync();
    }

    /// <summary>
    /// Add activity settings to an existing connection.
    /// </summary>
    /// <param name="connectionId">The connection ID of the connection to update.</param>
    /// <param name="itemType">The item type for the connection (`issues` or `repos`).</param>
    /// <returns>A <see cref="Task"/> indicating the status of the asynchronous update operation.</returns>
    public Task AddActivitySettingsAsync(string connectionId, string itemType)
    {
        var update = new ExternalConnection
        {
            ActivitySettings = new()
            {
                UrlToItemResolvers = new()
                {
                    new ItemIdResolver
                    {
                        Priority = 1,
                        ItemId = itemType == "issues" ? "{issueId}" : "{repo}",
                        UrlMatchInfo = new()
                        {
                            UrlPattern = itemType == "issues" ?
                                $"/{gitHubOwner}/{gitHubRepo}/issues/(?<issueId>[0-9]+)" :
                                $"/{gitHubOwner}/(?<repo>.*)/",
                            BaseUrls = new() { "https://github.com" },
                        },
                    },
                },
            },
        };

        return graphClient.External.Connections[connectionId].PatchAsync(update);
    }

    /// <summary>
    /// Registers a schema for a connection.
    /// </summary>
    /// <param name="connectionId">The connection ID of the connection.</param>
    /// <param name="schema">The <see cref="Schema"/> to register.</param>
    /// <returns>A <see cref="Task"/> indicating the status of the asynchronous registration operation.</returns>
    /// <exception cref="Exception">Thrown if a native HTTP request cannot be created.</exception>
    /// <exception cref="ServiceException">Thrown if the HTTP POST request to register the schema fails.</exception>
    public async Task RegisterSchemaAsync(string connectionId, Schema schema)
    {
        // Use the Graph SDK's request builder to generate the request URL
        var requestInfo = graphClient.External
            .Connections[connectionId]
            .Schema
            .ToGetRequestInformation();

        // Add the schema to the request content
        requestInfo.SetContentFromParsable(graphClient.RequestAdapter, "application/json", schema);

        // Convert the request to an HttpRequestMessage
        var requestMessage = await graphClient.RequestAdapter
            .ConvertToNativeRequestAsync<HttpRequestMessage>(requestInfo);
        _ = requestMessage ?? throw new Exception("Could not create native HTTP request");

        // Change method to POST
        requestMessage.Method = HttpMethod.Post;

        // Set required Prefer header
        // requestMessage.Headers.Add("Prefer", "respond-async");

        // Send the request
        var responseMessage = await httpClient.SendAsync(requestMessage) ??
            throw new Exception("No response returned from API");

        if (responseMessage.IsSuccessStatusCode)
        {
            // The operation ID is contained in the Location header returned
            // in the response
            var operationId = responseMessage.Headers.Location?.Segments.Last() ??
                throw new Exception("Could not get operation ID from Location header");

            var cts = new CancellationTokenSource();
            cts.CancelAfter(TimeSpan.FromMinutes(25));
            await WaitForOperationToCompleteAsync(connectionId, operationId, cts.Token);
        }
        else
        {
            throw new ServiceException(
                "Registering schema failed",
                responseMessage.Headers,
                (int)responseMessage.StatusCode);
        }
    }

    /// <summary>
    /// Adds or updates an <see cref="ExternalItem"/>.
    /// </summary>
    /// <param name="connectionId">The connection ID of the connection that contains the item.</param>
    /// <param name="item">The item to add or update.</param>
    /// <returns>The item.</returns>
    public Task<ExternalItem?> AddOrUpdateItemAsync(string connectionId, ExternalItem item)
    {
        return graphClient.External
            .Connections[connectionId]
            .Items[item.Id]
            .PutAsync(item);
    }

    /// <summary>
    /// Adds activities to an existing item.
    /// </summary>
    /// <param name="connectionId">The connection ID of the connection that contains the item.</param>
    /// <param name="itemId">The item ID of the item to update.</param>
    /// <param name="activities">The list of activities to add to the item.</param>
    /// <returns>The <see cref="AddActivitiesResponse"/>.</returns>
    public Task<AddActivitiesResponse?> AddIssueActivitiesAsync(string connectionId, string itemId, List<ExternalActivity> activities)
    {
        var addActivitiesRequest = new AddActivitiesPostRequestBody
        {
            Activities = activities,
        };

        return graphClient.External
            .Connections[connectionId]
            .Items[itemId]
            .MicrosoftGraphExternalConnectorsAddActivities
            .PostAsync(addActivitiesRequest);
    }

    /// <summary>
    /// Gets an <see cref="Identity"/> from a GitHub login.
    /// </summary>
    /// <param name="gitHubLogin">The GitHub login to look up.</param>
    /// <returns>The <see cref="Identity"/> that corresponds to the GitHub login.</returns>
    public Task<Identity> GetIdentityForGitHubUserAsync(string gitHubLogin)
    {
        _ = gitHubLogin;

        return Task.FromResult(new Identity
        {
            Type = IdentityType.User,
            Id = placeholderUserId,
        });
    }

    private async Task WaitForOperationToCompleteAsync(
        string connectionId, string operationId, CancellationToken cancellationToken)
    {
        do
        {
            // Periodically check for completion
            var operation = await graphClient.External
                .Connections[connectionId]
                .Operations[operationId]
                .GetAsync(cancellationToken: cancellationToken) ?? throw new Exception("Could not get operation from API.");

            if (operation.Status == ConnectionOperationStatus.Completed)
            {
                return;
            }

            if (operation.Status == ConnectionOperationStatus.Failed)
            {
                throw new ServiceException(operation.Error?.Message ?? "Registering schema failed");
            }

            if (cancellationToken.IsCancellationRequested)
            {
                throw new ServiceException("Schema registration timed out while checking for status.");
            }

            await Task.Delay(5000, cancellationToken);
        }
        while (true);
    }
}
