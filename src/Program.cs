// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.CommandLine;
using GitHubConnector;
using GitHubConnector.Extensions;
using GitHubConnector.Services;
using Microsoft.Graph;
using Microsoft.Graph.Beta.Models.ExternalConnectors;
using Microsoft.Graph.Beta.Models.ODataErrors;
using Octokit;

var useTeamsAppConfigOption = new Option<bool>(new[] { "--use-teams-app-config", "-t", })
{
    Description = "Run the connector in Teams app config mode.",
    IsRequired = false,
};

var command = new RootCommand();
command.AddOption(useTeamsAppConfigOption);

command.SetHandler(async (context) =>
{
    var useTeamsAppConfig = context.ParseResult.GetValueForOption(useTeamsAppConfigOption);

    try
    {
        var settings = AppSettings.LoadSettings();

        var connectorService = new SearchConnectorService(settings);
        var repoService = new RepositoryService(settings);

        if (useTeamsAppConfig)
        {
            await ListenForTeamsAppConfigAsync(connectorService, repoService);
        }
        else
        {
            await RunInteractivelyAsync(connectorService, repoService);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
});

Environment.Exit(await command.InvokeAsync(args));

static async Task ListenForTeamsAppConfigAsync(
    SearchConnectorService connectorService, RepositoryService repoService)
{
    await Task.Delay(1000);
    throw new NotImplementedException();
}

static async Task RunInteractivelyAsync(
    SearchConnectorService connectorService, RepositoryService repoService)
{
    ExternalConnection? currentConnection = null;
    try
    {
        do
        {
            var choice = Interactive.DoMenuPrompt(currentConnection);

            switch (choice)
            {
                case MenuChoice.CreateConnection:
                    var newConnection = await CreateConnectionInteractivelyAsync(connectorService);
                    currentConnection = newConnection ?? currentConnection;
                    break;
                case MenuChoice.SelectConnection:
                    var selectedConnection = await SelectConnectionInteractivelyAsync(connectorService);
                    currentConnection = selectedConnection ?? currentConnection;
                    break;
                case MenuChoice.DeleteConnection:
                    if (currentConnection != null)
                    {
                        await DeleteConnectionInteractivelyAsync(connectorService, currentConnection.Id);
                        currentConnection = null;
                    }
                    else
                    {
                        Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
                    }

                    break;
                case MenuChoice.RegisterSchema:
                    if (currentConnection != null)
                    {
                        await RegisterSchemaInteractivelyAsync(connectorService, currentConnection.Id);
                    }
                    else
                    {
                        Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
                    }

                    break;
                case MenuChoice.PushAllItems:
                    if (currentConnection != null)
                    {
                        await PushItemsInteractivelyAsync(connectorService, repoService, currentConnection.Id);
                    }
                    else
                    {
                        Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
                    }

                    break;
                case MenuChoice.Exit:
                    return;
                default:
                    Console.WriteLine("Invalid choice!");
                    break;
            }
        }
        while (true);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"ERROR: {ex.Message}");
    }
}

static async Task<ExternalConnection?> CreateConnectionInteractivelyAsync(SearchConnectorService connectorService)
{
    var connectionId = Interactive.PromptForInput(
        "Enter a unique ID for the new connection (3-32 characters)", true) ?? string.Empty;
    var connectionName = Interactive.PromptForInput(
        "Enter a name for the new connection", true) ?? string.Empty;
    var connectionDescription = Interactive.PromptForInput(
        "Enter a description for the new connection", false);
    var itemType = Interactive.PromptForItemType();

    try
    {
        var connection = await connectorService.CreateConnectionAsync(
            connectionId, connectionName, connectionDescription, itemType);

        Console.WriteLine($"New connection created - Name: {connection?.Name}, Id: {connection?.Id}");
        return connection;
    }
    catch (ODataError oDataError)
    {
        Console.WriteLine($"Error creating connection: {oDataError.ResponseStatusCode}");
        Console.WriteLine($"{oDataError.Error?.Code}: {oDataError.Error?.Message}");
        return null;
    }
}

static async Task<ExternalConnection?> SelectConnectionInteractivelyAsync(SearchConnectorService connectorService)
{
    Console.WriteLine("Getting existing connections...");
    try
    {
        var response = await connectorService.GetConnectionsAsync();
        var connections = response?.Value ?? new List<ExternalConnection>();
        if (connections.Count <= 0)
        {
            Console.WriteLine("No connections exist. Please create a new connection.");
            return null;
        }

        var selectedIndex = Interactive.PromptToSelectConnection(connections);

        return connections[selectedIndex];
    }
    catch (ODataError oDataError)
    {
        Console.WriteLine($"Error getting connections: {oDataError.ResponseStatusCode}");
        Console.WriteLine($"{oDataError.Error?.Code}: {oDataError.Error?.Message}");
        return null;
    }
}

static async Task DeleteConnectionInteractivelyAsync(SearchConnectorService connectorService, string? connectionId)
{
    _ = connectionId ?? throw new ArgumentNullException(nameof(connectionId));

    try
    {
        await connectorService.DeleteConnectionAsync(connectionId);
        Console.WriteLine("Connection deleted successfully.");
    }
    catch (ODataError oDataError)
    {
        Console.WriteLine($"Error deleting connection: {oDataError.ResponseStatusCode}");
        Console.WriteLine($"{oDataError.Error?.Code}: {oDataError.Error?.Message}");
    }
}

static async Task RegisterSchemaInteractivelyAsync(SearchConnectorService connectorService, string? connectionId)
{
    _ = connectionId ?? throw new ArgumentNullException(nameof(connectionId));

    var itemType = Interactive.PromptForItemType();
    Console.WriteLine("Registering schema, this may take some time...");

    try
    {
        await connectorService.RegisterSchemaAsync(
            connectionId,
            itemType == "issues" ? SearchConnectorService.IssuesSchema : SearchConnectorService.ReposSchema);
        Console.WriteLine("Schema registered successfully.");
    }
    catch (ODataError oDataError)
    {
        Console.WriteLine($"Error registering schema: {oDataError.ResponseStatusCode}");
        Console.WriteLine($"{oDataError.Error?.Code}: {oDataError.Error?.Message}");
    }
    catch (ServiceException ex)
    {
        Console.WriteLine($"Error registering schema: {ex.ResponseStatusCode}");
        Console.WriteLine($"{ex.Message}");
    }
}

static async Task PushItemsInteractivelyAsync(
    SearchConnectorService connectorService, RepositoryService repoService, string? connectionId)
{
    _ = connectionId ?? throw new ArgumentNullException(nameof(connectionId));

    var itemType = Interactive.PromptForItemType();

    if (itemType == "issues")
    {
        await PushAllIssuesWithActivitiesAsync(connectorService, repoService, connectionId);
    }
    else
    {
        await PushAllRepositoriesAsync(connectorService, repoService, connectionId);
    }
}

static async Task PushAllIssuesWithActivitiesAsync(
    SearchConnectorService connectorService, RepositoryService repoService, string connectionId)
{
    IReadOnlyList<Issue>? issues = null;
    try
    {
        issues = await repoService.GetIssuesForRepositoryAsync();
    }
    catch (ApiException ex)
    {
        Console.WriteLine($"Error getting issues: {ex.Message}");
    }

    using var httpClient = new HttpClient();
    foreach (var issue in issues ?? new List<Issue>())
    {
        var events = await repoService.GetEventsForIssueWithRetryAsync(issue.Number, 3);
        var issueItem = await issue.ToExternalItem(events, connectorService);

        var response = await httpClient.GetAsync(issue.HtmlUrl);
        if (response.IsSuccessStatusCode)
        {
            issueItem.Content = new()
            {
                Type = ExternalItemContentType.Html,
                Value = await response.Content.ReadAsStringAsync(),
            };
        }

        try
        {
            Console.Write($"Adding/updating issue {issue.Number}...");
            await connectorService.AddOrUpdateItemAsync(connectionId, issueItem);

            // Add activities from timeline events
            var activities = await events.ToExternalActivityList(connectorService);
            await connectorService.AddIssueActivitiesAsync(connectionId, issue.Number.ToString(), activities);
            Console.WriteLine("DONE");
        }
        catch (ODataError oDataError)
        {
            Console.WriteLine($"Error adding/updating issue: {oDataError.ResponseStatusCode}");
            Console.WriteLine($"{oDataError.Error?.Code}: {oDataError.Error?.Message}");
        }
    }
}

static async Task PushAllRepositoriesAsync(
    SearchConnectorService connectorService, RepositoryService repoService, string connectionId)
{
    IReadOnlyList<Repository>? repositories = null;

    try
    {
        repositories = await repoService.GetRepositoriesAsync();
    }
    catch (ApiException ex)
    {
        Console.WriteLine($"Error getting repositories: {ex.Message}");
    }

    using var httpClient = new HttpClient();
    foreach (var repository in repositories ?? new List<Repository>())
    {
        var events = await repoService.GetEventsForRepoWithRetryAsync(repository.Id, 3);
        var repoItem = await repository.ToExternalItem(events, connectorService);

        if (repository.Visibility == RepositoryVisibility.Public)
        {
            // Get the HTML content for the repo
            var response = await httpClient.GetAsync(repository.HtmlUrl);
            if (response.IsSuccessStatusCode)
            {
                repoItem.Content = new()
                {
                    Type = ExternalItemContentType.Html,
                    Value = await response.Content.ReadAsStringAsync(),
                };
            }
        }
        else
        {
            // Set content to the JSON representation
            repoItem.Content = new()
            {
                Type = ExternalItemContentType.Text,
                Value = System.Text.Json.JsonSerializer.Serialize(repository),
            };
        }

        try
        {
            Console.Write($"Adding/updating repository {repository.Name}...");
            await connectorService.AddOrUpdateItemAsync(connectionId, repoItem);
            Console.WriteLine("DONE");
        }
        catch (ODataError oDataError)
        {
            Console.WriteLine($"Error adding/updating repository: {oDataError.ResponseStatusCode}");
            Console.WriteLine($"{oDataError.Error?.Code}: {oDataError.Error?.Message}");
        }
    }
}
