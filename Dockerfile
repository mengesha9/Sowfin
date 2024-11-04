# Build environment
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build-env

WORKDIR /App

# Expose port 
EXPOSE 80

# Copy everything
COPY . ./
# Restore as distinct layers
RUN dotnet restore
# Build and publish a release of the application
RUN dotnet publish -c Release -o out

# Build runtime image with a generic tag (no specific architecture)
FROM mcr.microsoft.com/dotnet/aspnet:8.0 


WORKDIR /App

# Copy compiled output from build stage
COPY --from=build-env /App/out .

# Set the correct ENTRYPOINT
ENTRYPOINT ["dotnet", "Sowfin.API.dll"]
