APP_NAME := aseet
PKG := ./...
BIN_DIR := bin

.PHONY: all tidy lint test build license version

all: tidy lint test build license version

tidy:
	go mod tidy

lint:
	go vet $(PKG)

test:
	go test $(PKG) -cover

build:
	go build -ldflags "-s -w" -trimpath -o $(BIN_DIR)/$(APP_NAME).exe

license:
	go-licenses report $(PKG) --template=licenses.tpl --ignore $(APP_NAME) > NOTICE.md

version:
	$(BIN_DIR)/$(APP_NAME).exe version
