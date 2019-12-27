using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AutoMapper;
using DotNETCoreAPIMapExcelToObject.Models;
using DotNETCoreAPIMapExcelToObject.Ninject;
using ExcelToObject;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewFeatures.Internal;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.OpenApi.Models;
using Ninject;
using Ninject.Activation;
using Ninject.Infrastructure.Disposal;

namespace DotNETCoreAPIMapExcelToObject
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
            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);

            services.AddAutoMapper(typeof(Startup));

            #region swagger

            // Register the Swagger generator, defining 1 or more Swagger documents
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo
                {
                    Version = "v1",
                    Title = "MapExcelToObject",
                    Description = File.ReadAllText("Opdrachtomschrijving.txt"),
                    Contact = new OpenApiContact
                    {
                        Name = "Martien de Jong",
                        Email = "info@martiendejong.nl",
                        Url = new Uri("https://martiendejong.nl"),
                    },
                    License = new OpenApiLicense
                    {
                        Name = "Use under GPL 3.0",
                        Url = new Uri("https://www.gnu.org/licenses/gpl-3.0.en.html"),
                    }
                });
            });

            #endregion

            #region ninject

            services.AddSingleton<IHttpContextAccessor, HttpContextAccessor>();
            services.AddRequestScopingMiddleware(() => scopeProvider.Value = new Scope());
            services.AddCustomControllerActivation(Resolve);
            services.AddCustomViewComponentActivation(Resolve);
            
            #endregion
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            #region ninject

            Kernel = RegisterApplicationComponents(app);

            #endregion

            #region swagger

            // Enable middleware to serve generated Swagger as a JSON endpoint.
            app.UseSwagger();

            // Enable middleware to serve swagger-ui (HTML, JS, CSS, etc.),
            // specifying the Swagger JSON endpoint.
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "MapExcelToObject");
                c.RoutePrefix = string.Empty;
            });

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseHsts();
            }

            #endregion

            app.UseHttpsRedirection();
            app.UseMvc();
        }

        #region ninject

        private readonly AsyncLocal<Scope> scopeProvider = new AsyncLocal<Scope>();
        private IKernel Kernel { get; set; }

        private object Resolve(Type type) => Kernel.Get(type);
        private object RequestScope(IContext context) => scopeProvider.Value;

        private sealed class Scope : DisposableObject { }

        private IKernel RegisterApplicationComponents(IApplicationBuilder app)
        {
            // IKernelConfiguration config = new KernelConfiguration();
            var kernel = new StandardKernel();

            // Register application services
            foreach (var ctrlType in app.GetControllerTypes())
            {
                kernel.Bind(ctrlType).ToSelf().InScope(RequestScope);
            }

            ConfigureBindings(kernel);

            // Cross-wire required framework services
            kernel.BindToMethod(app.GetRequestService<IViewBufferScope>);

            return kernel;
        }

        #endregion

        // Ninject bindings
        public void ConfigureBindings(IKernel kernel)
        {
            kernel.Bind<IMapper>().ToMethod(ctx => new Mapper(CreateMapperConfiguration())).InSingletonScope();
            kernel.Bind<IExcelToObject<MappedObject>>().To<ExcelToObjectMapper<MappedObject>>().InSingletonScope();
            kernel.Bind<IExcelToObject<AnotherMappedObject>>().To<ExcelToObjectConfigFileMapper<AnotherMappedObject>>().InSingletonScope();
        }

        // Automapper configuration
        private MapperConfiguration CreateMapperConfiguration()
        {
            var config = new MapperConfiguration(cfg =>
            {
                // Add all profiles in current assembly
                cfg.AddProfiles(new Profile[] 
                {
                    new DataRowToMappedObjectMapping()
                });
            });

            return config;
        }
    }
}
