using OneDriveTidy.App.Components;
using OneDriveTidy.Core.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Register OneDriveTidy Services
string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
string dbPath = Path.Combine(appData, "OneDriveTidy", "onedrive_index.db");
// Ensure directory exists
Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);

builder.Services.AddSingleton<DatabaseService>(sp => new DatabaseService(dbPath));
builder.Services.AddSingleton<GraphService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
