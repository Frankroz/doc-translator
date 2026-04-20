using MongoDB.Driver;
using LoveNotesAdmin.Services;
using Microsoft.AspNetCore.Authentication.Cookies;
using LoveNotesAdmin;

DotNetEnv.Env.Load();

var builder = WebApplication.CreateBuilder(args);

// --- 1. REGISTRO DE SERVICIOS ---

// Soporte para Razor Pages y Controllers (necesario para el AuthController)
builder.Services.AddRazorPages();
builder.Services.AddControllers();

// Configuración de Blazor Server Interactivo
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Configuración de MongoDB
var mongoUri = Environment.GetEnvironmentVariable("MONGO_URI");
if (string.IsNullOrEmpty(mongoUri))
{
    throw new Exception("La variable de entorno MONGO_URI no está configurada.");
}
builder.Services.AddSingleton<IMongoClient>(new MongoClient(mongoUri));

// Servicios de Lógica de Negocio
builder.Services.AddScoped<AuthService>();
builder.Services.AddScoped<QuoteService>();

// Servicios Requeridos para Autenticación en Blazor
builder.Services.AddHttpContextAccessor();
builder.Services.AddCascadingAuthenticationState();

// Configurar Autenticación por Cookies
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath = "/login";
        options.Cookie.Name = "LoveNotes_Auth";
        options.Cookie.HttpOnly = true;
        options.Cookie.SameSite = SameSiteMode.Lax; // Cambiado de Strict a Lax
        options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest; 
        options.ExpireTimeSpan = TimeSpan.FromDays(7);
    });

builder.Services.AddAuthorization();

var app = builder.Build();

// --- 2. CONFIGURACIÓN DEL PIPELINE (Middleware) ---

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

// EL ORDEN ES CRÍTICO AQUÍ:
app.UseAuthentication(); // 1. ¿Quién eres?
app.UseAuthorization();  // 2. ¿Tienes permiso?
app.UseAntiforgery();    // 3. Protección contra ataques CSRF

// --- 3. MAPEO DE ENDPOINTS ---

app.MapControllers(); // Importante para el AuthController (api/auth/login)
app.MapRazorPages();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();