# Use official node image as the base image
FROM node:14.18-alpine3.14 as build

# Set the working directory
WORKDIR /usr/local/app

# Add the source code to app
COPY ./ ./

# Install all the dependencies
RUN npm ci

# Generate the build of the application
RUN npm run build:prod


# Stage 2: Serve app with nginx server

# Use official nginx image as the base image
FROM nginx:latest

# Copy the build output to replace the default nginx contents.
COPY --from=build /usr/local/app/dist/ /usr/share/nginx/html
#COPY /nginx.conf  /etc/nginx/conf.d/default.conf
# Assign permisson 
# https://stackoverflow.com/questions/49254476/getting-forbidden-error-while-using-nginx-inside-docker
# RUN chown nginx:nginx /usr/share/nginx/html/*

# Expose port 80
EXPOSE 80
EXPOSE 443