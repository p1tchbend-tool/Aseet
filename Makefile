APP_NAME := aseet
PKG := ./...
BIN_DIR := bin

# 環境変数を子プロセス（go buildなど）にエクスポートする設定
export GOOS
export GOARCH

.PHONY: all tidy lint test build build-windows build-darwin license version

all: tidy lint test build license version

tidy:
	go mod tidy

lint:
	go vet $(PKG)

test:
	go test $(PKG) -cover

build: build-windows build-darwin

# Windows (amd64) 向けビルド
build-windows: GOOS := windows
build-windows: GOARCH := amd64
build-windows:
	go build -ldflags "-s -w" -trimpath -o $(BIN_DIR)/$(APP_NAME)_windows_amd64.exe

# macOS (arm64) 向けビルド
build-darwin: GOOS := darwin
build-darwin: GOARCH := arm64
build-darwin:
	go build -ldflags "-s -w" -trimpath -o $(BIN_DIR)/$(APP_NAME)_darwin_arm64

license:
	go-licenses report $(PKG) --template=licenses.tpl --ignore $(APP_NAME) > NOTICE.md

version:
	go run main.go version
