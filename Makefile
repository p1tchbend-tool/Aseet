APP_NAME := aseet
PKG := ./...
BIN_DIR := bin

.PHONY: all tidy lint test build build-windows build-darwin license version

all: tidy lint test build license version

tidy:
	go mod tidy

lint:
	go vet $(PKG)

test:
	go test $(PKG) -cover

build: build-windows build-darwin

build-windows:
    GOOS=windows GOARCH=amd64 go build -ldflags "-s -w" -trimpath -o $(BIN_DIR)/$(APP_NAME)_windows_amd64.exe

build-darwin:
    GOOS=darwin GOARCH=arm64 go build -ldflags "-s -w" -trimpath -o $(BIN_DIR)/$(APP_NAME)_darwin_arm64

license:
	go-licenses report $(PKG) --template=licenses.tpl --ignore $(APP_NAME) > NOTICE.md

version:
	go run main.go version
