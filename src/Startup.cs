﻿// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.IdentityModel.Tokens;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Hosting;
using Microsoft.OpenApi.Models;
using Prometheus;
using NATS.Client;

using openrmf_read_api.Models;
using openrmf_read_api.Data;
using NLog.Web;
using NLog;

namespace openrmf_read_api
{
    public class Startup
    {
        readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
        
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            NLog.LogManager.Configuration = new NLog.Config.XmlLoggingConfiguration("nlog.config");
            var logger = NLog.LogManager.Setup().LoadConfigurationFromAppSettings().GetCurrentClassLogger();
            // Register the database components
            if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("DBTYPE")) || Environment.GetEnvironmentVariable("DBTYPE").ToLower() == "mongo") {
                services.Configure<Settings>(options =>
                {
                    options.ConnectionString = Environment.GetEnvironmentVariable("DBCONNECTION");
                    options.Database = Environment.GetEnvironmentVariable("DB");
                });
            }
            
            // Create a new connection factory to create a connection.
            ConnectionFactory cf = new ConnectionFactory();
            
            // add the options for the server, reconnecting, and the handler events
            Options opts = ConnectionFactory.GetDefaultOptions();
            opts.MaxReconnect = -1;
            opts.ReconnectWait = 1000;
            opts.Name = "openrmf-api-read";
            opts.Url = Environment.GetEnvironmentVariable("NATSSERVERURL");
            opts.AsyncErrorEventHandler += (sender, events) =>
            {
                logger.Info("NATS client error. Server: {0}. Message: {1}. Subject: {2}", events.Conn.ConnectedUrl, events.Error, events.Subscription.Subject);
            };

            opts.ServerDiscoveredEventHandler += (sender, events) =>
            {
                logger.Info("A new server has joined the cluster: {0}", events.Conn.DiscoveredServers);
            };

            opts.ClosedEventHandler += (sender, events) =>
            {
                logger.Info("Connection Closed: {0}", events.Conn.ConnectedUrl);
            };

            opts.ReconnectedEventHandler += (sender, events) =>
            {
                logger.Info("Connection Reconnected: {0}", events.Conn.ConnectedUrl);
            };

            opts.DisconnectedEventHandler += (sender, events) =>
            {
                logger.Info("Connection Disconnected: {0}", events.Conn.ConnectedUrl);
            };
            
            // Creates a live connection to the NATS Server with the above options
            IConnection conn = cf.CreateConnection(opts);

            // setup the NATS server
            services.Configure<NATSServer>(options =>
            {
                options.connection = conn;
            });

            // add repositories
            services.AddTransient<IArtifactRepository, ArtifactRepository>();
            services.AddTransient<ISystemGroupRepository, SystemGroupRepository>();

            // Register the Swagger generator, defining one or more Swagger documents
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo { Title = "OpenRMF Read API", Version = "v1", 
                    Description = "The Read API that goes with the OpenRMF OSS Application",
                    Contact = new OpenApiContact
                    {
                        Name = "Dale Bingham",
                        Email = "dale.bingham@cingulara.com",
                        Url = new Uri("https://github.com/Cingulara/openrmf-api-read")
                    } });
            });

            string jwtAuthorityServer = "http://openrmf-keycloak:8080/auth/"; // this is by default the internal keycloak setup
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("JWTINTERNALAUTHORITY"))) {
                jwtAuthorityServer = Environment.GetEnvironmentVariable("JWTINTERNALAUTHORITY").ToLower();
            }
            // Validate the JWT token sent with the request to make sure it is right
            // use the internal jwtAuthority server so you do not have to go outside the container network
            // allow the JWTINTERNALAUTHORITY if you need it for setup or for K8s specifically
            services.AddAuthentication(options =>
            {
                options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
                options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
            }).AddJwtBearer(o =>
            {
                o.Authority = jwtAuthorityServer + "realms/openrmf";
                o.Audience = Environment.GetEnvironmentVariable("JWTCLIENT");
                o.IncludeErrorDetails = true;
                o.RequireHttpsMetadata = false;
                o.TokenValidationParameters = new TokenValidationParameters()
                {
                    ValidateAudience = false,
                    ValidateIssuerSigningKey = true,
                    ValidateIssuer = true,
                    ValidIssuer = Environment.GetEnvironmentVariable("JWTAUTHORITY").ToLower(),
                    ValidateLifetime = true
                };

                o.Events = new JwtBearerEvents()
                {
                    OnAuthenticationFailed = c =>
                    {
                        c.NoResult();
                        c.Response.StatusCode = 401;
                        c.Response.ContentType = "text/plain";

                        Console.WriteLine("openrmf-api-read JWT Error: " + c.Exception.ToString());

                        return c.Response.WriteAsync("The JWT validation with the server did not return correctly. Please check with your Application Administrator.");
                    }
                };
            });

            // setup the RBAC for this
            services.AddAuthorization(options =>
            {
                options.AddPolicy("Administrator", policy => policy.RequireRole("roles", "[Administrator]"));
                options.AddPolicy("Download", policy => policy.RequireRole("roles", "[Download]"));
                options.AddPolicy("Editor", policy => policy.RequireRole("roles", "[Editor]"));
                options.AddPolicy("Reader", policy => policy.RequireRole("roles", "[Reader]"));
                options.AddPolicy("Assessor", policy => policy.RequireRole("roles", "[Assessor]"));
            });

            // add the CORS setup
            services.AddCors(options =>
            {
                options.AddPolicy(name: MyAllowSpecificOrigins,
                    builder =>
                    {
                        builder.AllowAnyHeader().AllowAnyMethod().AllowAnyOrigin();
                    });
            });
            
            // add service for allowing caching of responses
            services.AddResponseCaching();

            services.AddControllers();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            // Custom Metrics to count requests for each endpoint and the method
            var counter = Metrics.CreateCounter("openrmf_read_api_path_counter", "Counts requests to OpenRMF endpoints", new CounterConfiguration
            {
                LabelNames = new[] { "method", "endpoint" }
            });
            app.Use((context, next) =>
            {
                counter.WithLabels(context.Request.Method, context.Request.Path).Inc();
                return next();
            });
            // Use the Prometheus middleware
            app.UseMetricServer();
            app.UseHttpMetrics();

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            // Enable middleware to serve generated Swagger as a JSON endpoint.
            app.UseSwagger();

            // Enable middleware to serve swagger-ui (HTML, JS, CSS, etc.), specifying the Swagger JSON endpoint.
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "OpenRMF Read API V1");
            });

            // allow response caching directives in the API Controllers
            app.UseResponseCaching();
            app.UseHttpsRedirection();
            app.UseRouting();
            app.UseCors(MyAllowSpecificOrigins);
            // this has to go here
            app.UseAuthentication();
            app.UseAuthorization();
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
