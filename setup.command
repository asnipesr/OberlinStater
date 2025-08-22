#!/bin/bash

# cd ~/Desktop

# echo "Cloning CSU Bball Stater..."

# git clone https://github.com/asnipesr/BasketballStaterCSU.git

# echo "Completed Cloning. Checking Stater folder on Desktop"

cd ~/Desktop/OberlinStater

echo "Installing required packages"

pip3 install pygame
pip3 install openpyxl
brew install fileicon

echo "Installed Required Packages"

cp ~/Desktop/OberlinStater/run.sh ~/Desktop/OberlinStater/run2.sh
mv ~/Desktop/OberlinStater/run2.sh ~/Desktop/OberlinStater/OberlinStater.command
mv ~/Desktop/OberlinStater/"OberlinStater.command" ~/Desktop/OberlinStater.command
chmod +x ~/Desktop/OberlinStater.command
fileicon set ~/Desktop/OberlinStater.command ~/Desktop/OberlinStater/goyeo.png