VERSION ?= 0.7
NAME ?= "openrmf-api-read"
AUTHOR ?= "Dale Bingham"
PORT_EXT ?= 8084
PORT_INT ?= 8084
NO_CACHE ?= true
DOCKERHUB_ACCOUNT ?= cingulara
  
.PHONY: build docker latest clean version dockerhub

build:  
	dotnet build

docker: 
	docker build -f Dockerfile -t $(NAME)\:$(VERSION) --no-cache=$(NO_CACHE) .

latest: 
	docker build -f Dockerfile -t $(NAME)\:latest --no-cache=$(NO_CACHE) .
	docker login -u ${DOCKERHUB_ACCOUNT}
	docker tag $(NAME)\:latest ${DOCKERHUB_ACCOUNT}\/$(NAME)\:latest
	docker push ${DOCKERHUB_ACCOUNT}\/$(NAME)\:latest
	docker logout

 
clean:
	@rm -f -r obj
	@rm -f -r bin

version:
	@echo ${VERSION}

dockerhub:
	docker login -u ${DOCKERHUB_ACCOUNT}
	docker tag $(NAME)\:$(VERSION) ${DOCKERHUB_ACCOUNT}\/$(NAME)\:$(VERSION)
	docker push ${DOCKERHUB_ACCOUNT}\/$(NAME)\:$(VERSION)
	docker logout

DEFAULT: build
