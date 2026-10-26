using System.Collections.Concurrent;
using System.Net;
using System.Threading.Channels;
using System.Globalization;
using System.Text.Json;

var builder = WebApplication.CreateBuilder(args);
var root = Environment.GetEnvironmentVariable("RTU_REPORT_ROOT") ?? Path.GetFullPath("reports");
var app = builder.Build();
var subscribers = new ConcurrentDictionary<Guid, Channel<bool>>();
app.Use(async (context, next) => {
    context.Response.Headers.CacheControl = "no-store";
    await next(context);
});
app.MapPost("/api/reports/refresh", async (HttpContext context) => {
    if (context.Connection.RemoteIpAddress is not { } ip || !IPAddress.IsLoopback(ip)) return Results.StatusCode(403);
    string? reportDate;
    try {
        using var body = await JsonDocument.ParseAsync(context.Request.Body, cancellationToken: context.RequestAborted);
        if (body.RootElement.ValueKind != JsonValueKind.Object ||
            !body.RootElement.TryGetProperty("reportDate", out var date) || date.ValueKind != JsonValueKind.String)
            return Results.Problem("reportDate is required as a yyyy-MM-dd string.", statusCode: 400);
        reportDate = date.GetString();
    } catch (JsonException) {
        return Results.Problem("Request body must be JSON containing reportDate.", statusCode: 400);
    }
    if (!DateOnly.TryParseExact(reportDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
        return Results.Problem("reportDate must be a valid yyyy-MM-dd date.", statusCode: 400);
    var folder = Path.Combine(root, reportDate!);
    if (!Directory.Exists(folder)) return Results.Problem("Report date folder does not exist.", statusCode: 404);
    try {
        foreach (var filename in new[] { "complete", "main.html", "no-reason.html", "summary.html" }) {
            var path = Path.Combine(folder, filename);
            if (!File.Exists(path)) return Results.Problem($"Missing report file or completion marker: {filename}", statusCode: 409);
            using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            if (filename != "complete" && file.Length == 0)
                return Results.Problem($"Report file is empty: {filename}", statusCode: 409);
        }
    } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
        app.Logger.LogWarning(error, "Unable to validate report {ReportDate}", reportDate);
        return Results.Problem("Report files cannot currently be read. Retry the notification later.", statusCode: 503);
    }
    foreach (var channel in subscribers.Values) channel.Writer.TryWrite(true);
    return Results.Ok(new { notified = true, reportDate });
});
app.MapGet("/api/reports/events", async (HttpContext context) => {
    var id = Guid.NewGuid();
    var channel = Channel.CreateBounded<bool>(new BoundedChannelOptions(1) { FullMode = BoundedChannelFullMode.DropOldest });
    subscribers[id] = channel;
    context.Response.ContentType = "text/event-stream";
    context.Response.Headers["X-Accel-Buffering"] = "no";
    try {
        await context.Response.WriteAsync("event: ready\ndata: {}\n\n", context.RequestAborted);
        await context.Response.Body.FlushAsync(context.RequestAborted);
        await foreach (var _ in channel.Reader.ReadAllAsync(context.RequestAborted)) {
            await context.Response.WriteAsync("event: reports-updated\ndata: {}\n\n", context.RequestAborted);
            await context.Response.Body.FlushAsync(context.RequestAborted);
        }
    } catch (OperationCanceledException) when (context.RequestAborted.IsCancellationRequested) {
    } finally {
        subscribers.TryRemove(id, out _);
        channel.Writer.TryComplete();
    }
});
string[] Dates() => Directory.Exists(root) ? Directory.GetDirectories(root).Select(Path.GetFileName).OfType<string>().Where(x => DateOnly.TryParseExact(x,"yyyy-MM-dd",out _) && File.Exists(Path.Combine(root,x,"complete"))).OrderDescending().Take(7).ToArray() : [];
app.MapGet("/api/reports", () => Dates());
app.MapMethods("/reports/{date}/{type}", new[] { "GET", "HEAD" }, (string date,string type) => {
 if (!Dates().Contains(date) || !new[]{"main","no-reason","summary"}.Contains(type)) return Results.NotFound();
 var file=Path.Combine(root,date,type+".html");
 return File.Exists(file) ? Results.File(file,"text/html; charset=utf-8") : Results.NotFound();
});
app.MapGet("/", () => Results.Content(File.ReadAllText(Path.Combine(app.Environment.ContentRootPath, "Pages", "index.html")), "text/html; charset=utf-8"));
app.Run();
