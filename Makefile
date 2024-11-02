# Makefile for Sowfin project

# Target to build the project
build:
	@dotnet build

# Target to run the project after building it
run: build
	@dotnet run --project Sowfin.API

# Target to run the project with watch enabled
watch:
	@dotnet watch --project Sowfin.API run

# Clean target to remove build artifacts
clean:
	@dotnet clean

# Target to run the migrations
migrate:
	@dotnet ef migrations add InitialMigrations --project Sowfin.Data --startup-project Sowfin.API --context FindataContext

# Target to update the database
update:
	@dotnet ef database update --project Sowfin.Data --startup-project Sowfin.API --context FindataContext

# Target to restore dependencies
restore:
	@dotnet restore

# Target to publish the project
publish:
	@dotnet publish --configuration Release --output ./publish

# Help target to display available commands
help:
	@echo "Makefile for .NET WebAPI project"
	@echo ""
	@echo "Available commands:"
	@echo "  build     - Build the project"
	@echo "  run       - Run the project (after building)"
	@echo "  watch     - Run the project with file watching enabled"
	@echo "  clean     - Clean the build artifacts"
	@echo "  migrate   - Run database migrations"
	@echo "  update    - Update the database"
	@echo "  restore   - Restore the project dependencies"
	@echo "  publish   - Publish the project to the ./publish directory"
	@echo "  help      - Display this help message"
