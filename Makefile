.PHONY: start

ENV ?= . $(shell pwd)/env/bin/activate; \
       PYTHONPATH=$(shell pwd)

start:
	$(ENV) fbs run
