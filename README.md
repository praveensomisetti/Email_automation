app - main file

# How to Deploy Streamlit app on EC2 instance

## 1. Login with your AWS console and launch an EC2 instance

## 2. Run the following commands

### Note: Do the port mapping to this port:- 8501

```bash
sudo apt update
```

```bash
sudo apt-get update
```

```bash
sudo apt upgrade -y
```

```bash
sudo apt install git curl unzip tar make sudo vim wget -y
```

```bash
sudo apt install git curl unzip tar make sudo vim wget -y
```

```bash
git clone "Your-repository"
```

```bash
sudo apt install python3-pip
```
Install the python3-venv package to enable the creation of virtual environments.
sudo apt-get install python3-venv
python3 -m venv myenv
source myenv/bin/activate

```bash
pip3 install -r requirements.txt
```
Create the .env File
you need to create a .env file in your project directory on the EC2 instance. This file will hold your environment variables.
Save and exit the file by pressing CTRL + X, then Y, and ENTER.

```bash
#Temporary running
python3 -m streamlit run app.py
```

```bash
#Permanent running
nohup python3 -m streamlit run app.py
```

Note: Streamlit runs on this port: 8501



