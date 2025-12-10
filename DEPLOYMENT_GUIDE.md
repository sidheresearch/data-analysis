# Excel Refractor - Flask Application

A Flask-based web application for processing Excel files with features for data cleaning, seller analysis, and seller comparison.

## Project Overview

This application provides:
- **Data Cleaner**: Clean and process Excel data
- **Seller Analysis**: Analyze seller information and GST data
- **Seller Comparison**: Compare seller data across files
- File upload and download capabilities

## Prerequisites

- Python 3.7 or higher
- Amazon Linux 2023 (or any Linux distribution)
- Git (for version control)

## Complete Setup Guide for EC2 Hosting

### Step 1: Push Project to GitHub (From Your Local Machine)

1. **Initialize Git repository** (if not already done):
```bash
cd c:\Users\AMITKUMAR\Downloads\refractor\refractor
git init
```

2. **Add all files**:
```bash
git add .
```

3. **Commit changes**:
```bash
git commit -m "Initial commit: Excel Refractor Flask application"
```

4. **Create a new repository on GitHub**:
   - Go to https://github.com and log in
   - Click "New repository" (+ icon in top right)
   - Name: `excel-refractor` (or your preferred name)
   - Make it Private or Public
   - Do NOT initialize with README, .gitignore, or license
   - Click "Create repository"

5. **Link your local repo to GitHub**:
```bash
git remote add origin https://github.com/YOUR_USERNAME/excel-refractor.git
git branch -M main
git push -u origin main
```

### Step 2: Connect to Your EC2 Instance

You're already connected via SSH. If you need to reconnect:
```bash
ssh -i your-key.pem ec2-user@your-ec2-public-ip
```

### Step 3: Install Required System Packages on EC2

```bash
# Update system
sudo dnf update -y

# Install Python 3 and pip
sudo dnf install -y python3 python3-pip git

# Verify installation
python3 --version
git --version
```

### Step 4: Clone Your Project from GitHub

```bash
# Navigate to home directory
cd ~

# Clone your repository (replace with your GitHub URL)
git clone https://github.com/YOUR_USERNAME/excel-refractor.git

# Rename folder to 'refractor' if needed
mv excel-refractor refractor

# Navigate to project
cd refractor
```

### Step 5: Set Up Virtual Environment and Install Dependencies

```bash
# Make setup script executable
chmod +x setup.sh

# Run setup script
bash setup.sh
```

This will:
- Create a Python virtual environment (`venv`)
- Install all required packages from `requirements.txt`
- Set up the project structure

### Step 6: Test the Application (Development Mode)

```bash
# Make run script executable
chmod +x run.sh

# Run the application
bash run.sh
```

The app will run on `http://0.0.0.0:5000`

**Important**: Configure EC2 Security Group to allow incoming traffic on port 5000:
1. Go to EC2 Console → Security Groups
2. Select your instance's security group
3. Add inbound rule: Type=Custom TCP, Port=5000, Source=0.0.0.0/0 (or your IP)

Access the app at: `http://YOUR_EC2_PUBLIC_IP:5000`

Press `Ctrl+C` to stop the development server.

### Step 7: Set Up Production Service (Run in Background)

For production, use systemd to run the app as a background service:

```bash
# Edit the service file if needed (update paths if different)
nano refractor.service

# Copy service file to systemd
sudo cp refractor.service /etc/systemd/system/

# Reload systemd
sudo systemctl daemon-reload

# Enable service to start on boot
sudo systemctl enable refractor.service

# Start the service
sudo systemctl start refractor.service

# Check status
sudo systemctl status refractor.service
```

**Service Management Commands**:
```bash
# Start service
sudo systemctl start refractor.service

# Stop service
sudo systemctl stop refractor.service

# Restart service
sudo systemctl restart refractor.service

# View logs
sudo journalctl -u refractor.service -f
```

### Step 8: Optional - Set Up Nginx as Reverse Proxy

For production use with a domain name or better security:

```bash
# Install Nginx
sudo dnf install -y nginx

# Create Nginx configuration
sudo nano /etc/nginx/conf.d/refractor.conf
```

Add this configuration:
```nginx
server {
    listen 80;
    server_name your-domain.com;  # or your EC2 public IP

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

```bash
# Start and enable Nginx
sudo systemctl start nginx
sudo systemctl enable nginx

# Update Security Group to allow HTTP (port 80)
```

Now access your app at: `http://YOUR_EC2_PUBLIC_IP` (port 80)

### Step 9: Update Security Group Settings

**For Development (Port 5000)**:
- Type: Custom TCP
- Port: 5000
- Source: 0.0.0.0/0 (or your specific IP)

**For Production with Nginx (Port 80)**:
- Type: HTTP
- Port: 80
- Source: 0.0.0.0/0

**For HTTPS (Recommended)**:
- Type: HTTPS
- Port: 443
- Source: 0.0.0.0/0

### Step 10: Updating Your Application

When you make changes locally:

```bash
# On local machine
git add .
git commit -m "Description of changes"
git push origin main

# On EC2 instance
cd ~/refractor
git pull origin main

# If you updated dependencies
source venv/bin/activate
pip install -r requirements.txt

# Restart the service
sudo systemctl restart refractor.service
```

## Project Structure

```
refractor/
├── app.py                  # Main Flask application
├── check_price_file.py     # Price checking utility
├── requirements.txt        # Python dependencies
├── setup.sh               # Setup script for Linux
├── run.sh                 # Run script for development
├── run.bat                # Run script for Windows
├── refractor.service      # Systemd service file
├── .gitignore            # Git ignore rules
├── templates/            # HTML templates
│   ├── index.html
│   ├── data_cleaner.html
│   ├── seller_analysis.html
│   └── ...
├── uploads/              # Uploaded files (not in Git)
├── processed/            # Processed files (not in Git)
└── cache/                # Session cache (not in Git)
```

## Important Security Notes

1. **Change the secret key** in `app.py`:
   ```python
   app.secret_key = 'your-secret-key-here-change-this-to-random-string'
   ```
   Generate a random one:
   ```bash
   python3 -c "import secrets; print(secrets.token_hex(32))"
   ```

2. **Set up HTTPS** for production using Let's Encrypt:
   ```bash
   sudo dnf install -y certbot python3-certbot-nginx
   sudo certbot --nginx -d your-domain.com
   ```

3. **Use environment variables** for sensitive data

4. **Restrict Security Group** access to specific IPs when possible

## Troubleshooting

### Application won't start
```bash
# Check service status
sudo systemctl status refractor.service

# View logs
sudo journalctl -u refractor.service -n 50
```

### Port already in use
```bash
# Find process using port 5000
sudo lsof -i :5000

# Kill process
sudo kill -9 PID
```

### Permission issues
```bash
# Ensure proper ownership
sudo chown -R ec2-user:ec2-user ~/refractor
```

### Virtual environment issues
```bash
# Remove and recreate
rm -rf venv
bash setup.sh
```

## Multiple Projects on Same Instance

To host multiple projects:

1. **Use different directories**:
   ```bash
   ~/project1/
   ~/project2/
   ```

2. **Use different ports** in each app (5000, 5001, etc.)

3. **Create separate virtual environments** for each

4. **Create separate systemd services** for each

5. **Use Nginx** to route traffic based on domain/path:
   ```nginx
   # Project 1
   server {
       listen 80;
       server_name project1.com;
       location / {
           proxy_pass http://127.0.0.1:5000;
       }
   }
   
   # Project 2
   server {
       listen 80;
       server_name project2.com;
       location / {
           proxy_pass http://127.0.0.1:5001;
       }
   }
   ```

## Support

For issues or questions, check the logs:
```bash
# Application logs
sudo journalctl -u refractor.service -f

# Nginx logs
sudo tail -f /var/log/nginx/error.log
sudo tail -f /var/log/nginx/access.log
```

## License

Add your license information here.
