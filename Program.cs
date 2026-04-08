using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using gAcss.Models;
using DriveAnalytic.Services;
using gAcss.Service;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.AddSingleton<ExcelExportService>();
builder.Services.AddSingleton<IDriveService, GoogleDriveProvider>();
builder.Services.AddSingleton<AppRunner>();

// Регистрация Google Client
// Регистрация Google Client через нашу фабрику
builder.Services.AddSingleton<DriveService>(sp =>
{
    // Используем .GetAwaiter().GetResult(), так как это Singleton, 
    // который создается один раз при первом обращении.
    return GoogleClientFactory.CreateDriveServiceAsync().GetAwaiter().GetResult();
});

using IHost host = builder.Build();

var app = host.Services.GetRequiredService<AppRunner>();
await app.RunAsync();