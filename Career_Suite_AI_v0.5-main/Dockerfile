# Use an official Node.js runtime as a parent image
FROM node:18

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy package.json and package-lock.json (if available)
COPY package*.json ./

# Install app dependencies using npm ci for potentially faster/cleaner installs if lock file exists
# Using npm install is also fine
RUN npm install

# Bundle app source inside Docker image
COPY . .

# Define the command to run your app using CMD which defines your runtime
CMD ["npm", "start"]