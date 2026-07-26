# OfficeEditor development commands
# Run `make dev` to start the API and React client side by side.

API_DIR      := OfficeEditor.Api
WEB_DIR      := OfficeEditor.Web.Client
API_URL      := http://localhost:5001
WEB_PORT     := 5173
DOTNET_ROLL  := DOTNET_ROLL_FORWARD=Major

.PHONY: dev api web build test clean install help

## Run the API with Blazor Server demo (single process, no Node needed)
dev:
	cd $(API_DIR) && $(DOTNET_ROLL) DOTNET_ENVIRONMENT=Development dotnet run --no-launch-profile --urls "$(API_URL)"

## Run only the ASP.NET Core API
api:
	cd $(API_DIR) && $(DOTNET_ROLL) DOTNET_ENVIRONMENT=Development dotnet run --no-launch-profile --urls "$(API_URL)"

## Run only the React client (expects the API to be running on $(API_URL))
web:
	cd $(WEB_DIR) && npm run dev

## Build the .NET solution and the React client
build:
	dotnet build DocxEditor.sln
	cd $(WEB_DIR) && npm run build

## Run all .NET tests and verify the React client builds
test:
	dotnet test DocxEditor.sln
	cd $(WEB_DIR) && npm run build

## Install / restore frontend dependencies
install:
	cd $(WEB_DIR) && npm install

## Clean .NET build output and the React dist folder
clean:
	dotnet clean DocxEditor.sln
	cd $(WEB_DIR) && rm -rf dist

## Show available commands
help:
	@echo "Available commands:"
	@echo "  make dev     - Start API + React client side by side"
	@echo "  make api     - Start only the API"
	@echo "  make web     - Start only the React client"
	@echo "  make build   - Build .NET solution + React client"
	@echo "  make test    - Run .NET tests + React build check"
	@echo "  make install - Install React dependencies"
	@echo "  make clean   - Clean build outputs"
