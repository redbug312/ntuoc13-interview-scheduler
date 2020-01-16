.PHONY: start


ifeq ($(OS), Windows_NT)
    ENV ?= . $(shell pwd)/env/scripts/activate; \
        PYTHONPATH=$(shell pwd) \
        PATH=/c/Program\ Files\ \(x86\)/NSIS/:$$PATH
else
    ENV ?= . $(shell pwd)/env/bin/activate; \
        PYTHONPATH=$(shell pwd)
endif


all: start

build:
	$(ENV) fbs freeze
	$(ENV) fbs installer

start:
	$(ENV) fbs run
