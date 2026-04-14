using MongoDB.Driver;
using LoveNotesApi.Models;
using Microsoft.AspNetCore.Mvc;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

DotNetEnv.Env.Load();

// MongoDB Setup
// In a real app, move this connection string to appsettings.json
var mongoClient = new MongoClient(Environment.GetEnvironmentVariable("MONGO_URI"));
var database = mongoClient.GetDatabase("Translator");
var _quotesCollection = database.GetCollection<Quote>("quotes");

// Endpoint: Get Random Loading Quote
app.MapGet("/quote/loading", async () =>
{
    var quotes = await _quotesCollection.Find(q => q.Category == "loading").ToListAsync();
    if (!quotes.Any()) return Results.Ok(new { text = "Traduciendo con amor..." });

    var random = new Random();
    var selected = quotes[random.Next(quotes.Count)];
    return Results.Ok(new { text = selected.Text });
});

// Endpoint: Get Random Ending Quote
app.MapGet("/quote/ending", async () =>
{
    var quotes = await _quotesCollection.Find(q => q.Category == "ending").ToListAsync();
    if (!quotes.Any()) return Results.Ok(new { text = "¡Todo listo! ✨" });

    var random = new Random();
    var selected = quotes[random.Next(quotes.Count)];
    return Results.Ok(new { text = selected.Text });
});

app.Run();