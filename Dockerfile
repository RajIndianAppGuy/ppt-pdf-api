# Use an official Python runtime as a parent image
FROM python:3.12-slim

# Set environment variables to avoid buffering
ENV PYTHONUNBUFFERED=1

# Install LibreOffice and other dependencies
RUN apt-get update && apt-get upgrade -y
# Install Ghostscript
RUN apt-get install -y libreoffice 
# Update the package
RUN apt-get update
# Set the working directory
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install any needed packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Make port 5000 available to the world outside this container
EXPOSE 5000

# Run app.py when the container launches
CMD ["python", "app.py"]
