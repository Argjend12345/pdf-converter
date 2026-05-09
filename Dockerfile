# First Stage: Build the application
FROM golang:1.23.4-alpine3.21 AS builder

LABEL maintainer="wteja"

WORKDIR /app

COPY go.mod ./

RUN go mod download

COPY . .

RUN go build -o pdf-converter .

# Second Stage: Copy the binary and required files to a new image
FROM ubuntu:latest AS runner

ENV GDK_DPI_SCALE=1
ENV GDK_SCALE=1
ENV DEBIAN_FRONTEND=noninteractive

WORKDIR /app

RUN apt-get update && apt-get install -y libreoffice fonts-thai-tlwg ttf-mscorefonts-installer fontconfig && fc-cache -fv

COPY fonts /usr/share/fonts/custom

RUN fc-cache -fv

COPY --from=builder /app/pdf-converter /app/pdf-converter

EXPOSE 5000

VOLUME [ "/app/tmp" ]

CMD ["/app/pdf-converter"]