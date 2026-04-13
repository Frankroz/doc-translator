using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace LoveNotesApi.Models;

public class Quote
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string? Id { get; set; }

    [BsonElement("text")]
    public string Text { get; set; } = null!;

    [BsonElement("category")]
    public string Category { get; set; } = null!; // "loading" or "ending"
}