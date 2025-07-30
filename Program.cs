
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Extensions.DependencyInjection;
using System.Net.Http;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<BlazorOfficeAddinPOC.App>("#app");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new System.Uri(builder.HostEnvironment.BaseAddress) });

await builder.Build().RunAsync();