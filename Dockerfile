FROM node:18

# Install Python
RUN apt-get update && apt-get install -y python3 python3-pip && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Node.js dependencies
COPY package*.json ./
RUN npm install

# Install Python dependencies one by one (for better error tracking)
RUN pip3 install pandas==2.2.3
RUN pip3 install openpyxl==3.1.5

# Copy application files
COPY . .

EXPOSE 5000
CMD ["node", "server.js"]
