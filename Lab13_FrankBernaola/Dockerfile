# Utiliza imagen oficial de .NET 8 SDK
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 80

# Copia los archivos publicados desde la fase build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["Lab13_FrankBernaola/Lab13_FrankBernaola.csproj", "Lab13_FrankBernaola/"]
RUN dotnet restore "Lab13_FrankBernaola/Lab13_FrankBernaola.csproj"
COPY . .
WORKDIR "/src/Lab13_FrankBernaola"
RUN dotnet publish "Lab13_FrankBernaola.csproj" -c Release -o /app/publish

# Imagen final
FROM base AS final
WORKDIR /app
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "Lab13_FrankBernaola.dll"]
