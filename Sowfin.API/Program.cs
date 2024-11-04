using System.Xml.Schema;
using System.Net.Security;
using System.Net.Mime;
using System;
using System.Collections.Immutable;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using System.Reflection;
using Microsoft.OpenApi.Models;
using Sowfin.API.Notifications;
using Sowfin.API.Services;
using Sowfin.API.Services.Abstraction;
using Microsoft.AspNetCore.SignalR;

using Sowfin.API.ViewModels.Mapping;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Serialization;
using Npgsql;
using Sofcan.Lib;
using Sowfin.Data.Abstract;
using Sowfin.Data.Repositories;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using Sowfin.Data;
using Sowfin.API.Lib;


var builder = WebApplication.CreateBuilder(args);


builder.Services.AddControllers();
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", builder =>
    {
        builder.AllowAnyOrigin()
            .AllowAnyMethod()
            .AllowAnyHeader();
    });
});

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
            {
                var securitySchema = new OpenApiSecurityScheme
                {
                    Description = "JWT Auth Bearer Scheme",
                    Name = "Authorisation",
                    In = ParameterLocation.Header,
                    Type = SecuritySchemeType.Http,
                    Scheme = "Bearer",
                    Reference = new OpenApiReference
                    {
                        Type = ReferenceType.SecurityScheme,
                        Id = "Bearer"
                    }
                };

                c.AddSecurityDefinition("Bearer", securitySchema);

                var securityRequirement = new OpenApiSecurityRequirement
                {
                    {
                        securitySchema, new[] {"Bearer"}
                    }
                };

                c.AddSecurityRequirement(securityRequirement);

            });

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

builder.Services.AddHttpContextAccessor();
builder.Services.AddHttpClient();
builder.Services.AddAuthorization();
// builder.Services.AddScoped<IHttpContextAccessor, HttpContextAccessor>();
var configuration = builder.Configuration;


builder.Services.AddDbContext<FindataContext>((serviceProvider, options) =>
{
    var httpContext = serviceProvider.GetService<IHttpContextAccessor>().HttpContext;
    var cikValue = httpContext?.Request.Headers["Cik"];
    if (!string.IsNullOrEmpty(cikValue))
    {
        var connectionString = $"Server=localhost;Database=dataengine_{cikValue};Username=postgres;Password=system";
        NpgsqlConnectionStringBuilder connectionBuilder = new NpgsqlConnectionStringBuilder(connectionString);
        options.UseNpgsql(connectionBuilder.ToString(), o => o.MigrationsAssembly("Sowfin.Data"));
    }
    else
    {
        options.UseNpgsql(configuration.GetConnectionString("FindataContext"), o => o.MigrationsAssembly("Sowfin.Data"));
    }
});




builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = false,
            ValidateAudience = false,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,

            IssuerSigningKey = new SymmetricSecurityKey(
                Encoding.UTF8.GetBytes(configuration.GetValue<string>("JWTSecretKey"))
            )
        };

        options.Events = new JwtBearerEvents
        {
            OnMessageReceived = context =>
            {
                var accessToken = context.Request.Query["access_token"];

                // If the request is for our hub...
                var path = context.HttpContext.Request.Path;
                if (!string.IsNullOrEmpty(accessToken) && (path.StartsWithSegments("/notifications")))
                {
                    // Read the token out of the query string
                    context.Token = accessToken;
                }
                return Task.CompletedTask;
            }
        };
    });

builder.Services.AddAutoMapper(Assembly.GetExecutingAssembly());

builder.Services.AddDataConfigurationRegistration(configuration);


builder.Services.AddSingleton<IAuthService>(
                new AuthService(
                    configuration.GetValue<string>("JWTSecretKey"),
                    configuration.GetValue<int>("JWTLifespan")
                )
            );



builder.Services.AddSingleton<ISowfinCache>(provider =>
{

    return new SowfinCache(
        configuration.GetValue<string>("RedisServer"),
        configuration.GetValue<string>("RedisPort")
    );

}
);

builder.Services
    .AddControllers()
    .AddJsonOptions(options =>
    {
        // options.JsonSerializerOptions.ReferenceHandler = System.Text.Json.Serialization.ReferenceHandler.Preserve;


        options.JsonSerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;

        options.JsonSerializerOptions.Converters.Add(new JsonStringEnumConverter());
        options.JsonSerializerOptions.NumberHandling = JsonNumberHandling.AllowNamedFloatingPointLiterals;

    });




builder.Services.AddSignalR().AddJsonProtocol(options =>
{
    options.PayloadSerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
    options.PayloadSerializerOptions.Converters.Add(new JsonStringEnumConverter());
});


builder.Services.AddSingleton<IUserIdProvider, UserIdProvider>();


var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI();
// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();

}
else{
    app.UseHsts();

}

app.UseCors("AllowAll"); 

app.MapHub<NotificationsHub>("/notifications");


app.UseHttpsRedirection();
app.MapControllers();
app.UseAuthentication();
app.UseAuthorization();
app.UseRouting();
app.Run();



    // "FindataContext": "Server=localhost;Database=sowfin;Username=postgres;Password=nitesh"


