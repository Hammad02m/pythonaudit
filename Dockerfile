FROM node:18

# Install Python and pip
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

# Create symbolic link for python command
RUN ln -s /usr/bin/python3 /usr/bin/python

WORKDIR /app

# Copy and install Node.js dependencies first (for better caching)
COPY package*.json ./
RUN npm install

# Copy and install Python dependencies
COPY requirements.txt ./
RUN pip3 install -r requirements.txt

# Copy all application files
COPY . .

# Expose port 5000 (as specified in your server.js)
EXPOSE 5000

# Start your Node.js server
CMD ["node", "server.js"]
