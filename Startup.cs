using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Swashbuckle.AspNetCore.Swagger;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

using openstig_read_api.Models;
using openstig_read_api.Data;

namespace openstig_read_api
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            // Register the database components

            services.Configure<Settings>(options =>
            {
                options.ConnectionString = Environment.GetEnvironmentVariable("mongoConnection");
                options.Database = Environment.GetEnvironmentVariable("mongodb");
            });
            
            services.AddTransient<IArtifactRepository, ArtifactRepository>();

            // Register the Swagger generator, defining one or more Swagger documents
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new Info { Title = "openSTIG Scoring API", Version = "v1", 
                    Description = "The Scoring API that goes with the openSTIG tool",
                    Contact = new Contact
                    {
                        Name = "Dale Bingham",
                        Email = "dale.bingham@cingulara.com",
                        Url = "https://github.com/Cingulara/openstig-api-scoring"
                    } });
            });

            // ********************
            // USE CORS
            // ********************
            services.AddCors(options =>
            {
                options.AddPolicy("AllowAll",
                    builder =>
                    {
                        builder
                        .AllowAnyOrigin() 
                        .AllowAnyMethod()
                        .AllowAnyHeader()
                        .AllowCredentials();
                    });
            });
            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_2);

            // add this in memory for now. Persist later.
        	services.AddDistributedMemoryCache();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
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
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "openSTIG Score API V1");
            });

            // ********************
            // USE CORS
            // ********************
            app.UseCors("AllowAll");

            app.UseHttpsRedirection();
            app.UseMvc();
        }
    }
}
