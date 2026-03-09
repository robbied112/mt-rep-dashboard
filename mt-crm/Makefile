# Missing Thorn Dashboard Build System
#
# Usage:
#   make build                    # Build both dashboards (data from ./data/)
#   make build DATA_DIR=~/OneDrive/MT\ Dashboard\ Data/
#   make build-exec               # Build executive dashboard only
#   make build-rep                # Build rep dashboard only
#   make deploy                   # Deploy both to GitHub Pages
#   make deploy-exec              # Deploy executive only
#   make deploy-rep               # Deploy rep only
#   make all                      # Build + deploy everything

DATA_DIR ?= data/

.PHONY: all build build-exec build-rep deploy deploy-exec deploy-rep clean

all: build deploy

build: build-exec build-rep

build-exec:
	python3 scripts/build_dashboard.py \
		--data "$(DATA_DIR)" \
		--template templates/executive_dashboard.html \
		--output dist/index.html

build-rep:
	python3 scripts/build_rep_dashboard.py \
		--data "$(DATA_DIR)" \
		--template templates/rep_dashboard_template.html \
		--output dist/index_rep.html

deploy: deploy-exec deploy-rep

deploy-exec:
	./deploy.sh exec

deploy-rep:
	./deploy.sh rep

clean:
	rm -f dist/index.html dist/index_rep.html
