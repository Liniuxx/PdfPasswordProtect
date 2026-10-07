using VipaPDFProtect.Services;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddHttpClient();
builder.Services.AddSingleton<GraphTokenService>();
builder.Services.AddSingleton<GraphOneDriveService>();
builder.Services.AddSingleton<PasswordCryptoService>();
builder.Services.AddSingleton<PdfProtectService>();

var app = builder.Build();
app.UseSwagger();
app.UseSwaggerUI();
app.MapGet("/", () => Results.Ok(new { service = "VipaPDFProtect", status = "Running", endpoint = "POST /api/protect-pdf", swagger = "/swagger" }));
app.MapGet("/health", () => Results.Ok(new { status = "OK", utc = DateTime.UtcNow }));
app.MapControllers();
app.Run();
