using IBMWrapperAPI.Models;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.PlatformAbstractions;
using Swashbuckle.AspNetCore.Swagger;
using System.IO;

namespace IBMWrapperAPI
{
    /// <summary>
    /// Startup config
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// public constructor
        /// </summary>
        /// <param name="configuration"></param>
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        /// <summary>
        /// Config for startup
        /// </summary>
        public IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services"></param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddMvc();

            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new Info { Title = "IBM Wrapper API", Description = "Personality API" });
                
                var basePath = PlatformServices.Default.Application.ApplicationBasePath;
                var xmlPath = Path.Combine(basePath, "IBMWrapperAPI.xml");
                c.IncludeXmlComments(xmlPath);
                c.OperationFilter<FileUploadOperations>();
            });

            services.AddCors(c =>
            {
                c.AddPolicy("MyPolicy", p =>
                {
                    p.AllowAnyHeader();
                    p.AllowAnyOrigin();
                    p.AllowAnyMethod();
                });
            });
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app"></param>
        /// <param name="env"></param>
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseCors("MyPolicy");

            app.UseMvc(c =>
            {
                c.MapRoute("default", "{controller=Home}/{action=Index}");
            });

            app.UseSwagger();

            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "IBM Wrapper API");
            });
        }
    }
}
