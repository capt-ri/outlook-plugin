using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using OutlookEventPlugin;
using OutlookEventPlugin.Services;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });

// Mock Salesforce backend (in-memory). Swap for a real HTTP-backed impl later.
builder.Services.AddSingleton<ISalesforceSync, MockSalesforceSync>();

await builder.Build().RunAsync();
