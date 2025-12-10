#!/bin/bash
# Script to add swap space to improve performance

echo "Creating 2GB swap file..."
sudo dd if=/dev/zero of=/swapfile bs=1M count=2048
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile

# Make swap permanent
echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab

# Verify swap is active
echo "Swap space added successfully!"
free -h
