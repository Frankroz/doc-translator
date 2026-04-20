using LoveNotesAdmin.Models;
using MongoDB.Driver;
using Microsoft.AspNetCore.Identity;

namespace LoveNotesAdmin.Services
{
    public class QuoteService
    {
        private readonly IMongoCollection<Quote> _quotesCollection;

        public QuoteService(IMongoClient mongoClient, IConfiguration config)
        {
            
            var database = mongoClient.GetDatabase("Translator");
            _quotesCollection = database.GetCollection<Quote>("quotes");
        }

        public async Task<List<Quote>> GetAllAsync() =>
            await _quotesCollection.Find(_ => true).ToListAsync();

        public async Task CreateAsync(Quote newQuote) =>
            await _quotesCollection.InsertOneAsync(newQuote);

        public async Task UpdateAsync(string id, Quote updatedQuote) =>
            await _quotesCollection.ReplaceOneAsync(q => q.Id == id, updatedQuote);

        public async Task DeleteAsync(string id) =>
            await _quotesCollection.DeleteOneAsync(q => q.Id == id);
    }
}