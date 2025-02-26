using Application.Features.Vulnerability.Queries.FirstTask;
using Application.Interfaces;
using Infrastructure.Services;
using Microsoft.Extensions.DependencyInjection;

namespace Infrastructure;

public static class ServiceCollectionExtension
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddMediatR(cfg => cfg.RegisterServicesFromAssembly(typeof(VulnerabilityFirstTaskQuery).Assembly));
        services.AddScoped<IParser, Parser>();
        services.AddScoped<IAggregator, Aggregator>();
        services.AddScoped<IExcelChartExporter, ExcelChartExporter>();
        return services;
    }
}
