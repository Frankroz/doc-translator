using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace LoveNotesAdmin.Models
{
    public class Quote
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [BsonElement("text")]
        public string Text { get; set; } = string.Empty;

        [BsonElement("category")]
        public string Category { get; set; } = "loading"; // "loading" o "ending"
    }
}