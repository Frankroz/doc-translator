using LoveNotesAdmin.Models;
using MongoDB.Driver;
using Microsoft.AspNetCore.Identity;

namespace LoveNotesAdmin.Services;

public class AuthService
{
    private readonly IMongoCollection<AdminUser> _users;
    private readonly IPasswordHasher<AdminUser> _hasher;

    public AuthService(IMongoClient mongoClient, IConfiguration config)
    {
        var database = mongoClient.GetDatabase("Translator");
        _users = database.GetCollection<AdminUser>("admins");
        _hasher = new PasswordHasher<AdminUser>();
    }

    public async Task<AdminUser?> ValidateUser(string username, string password)
    {
        var user = await _users.Find(u => u.Username == username).FirstOrDefaultAsync();
        if (user == null) return null;

        // Comparamos el hash guardado con la contraseña ingresada
        var result = _hasher.VerifyHashedPassword(user, user.Password, password);
        return result == PasswordVerificationResult.Success ? user : null;
    }
}