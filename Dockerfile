# Use a small Python image
FROM python:3.11-slim

# Set working directory inside the container
WORKDIR /app

# Copy all files from your current folder into the container
COPY . /app

# Install Python dependencies listed in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Let IBM Code Engine know your app runs on port 8080
EXPOSE 8080

# Start the app using Python
CMD ["python", "app.py"]

