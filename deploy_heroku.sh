#!/bin/bash

echo "========================================"
echo "   DEPLOY ATTENDANCE SYSTEM TO HEROKU"
echo "========================================"
echo

echo "Checking if Heroku CLI is installed..."
if ! command -v heroku &> /dev/null; then
    echo "ERROR: Heroku CLI not found!"
    echo "Please install from: https://devcenter.heroku.com/articles/heroku-cli"
    exit 1
fi

echo "Heroku CLI found!"
echo

echo "Please login to Heroku:"
heroku login
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to login to Heroku"
    exit 1
fi

echo
read -p "Enter your Heroku app name (e.g., my-attendance-system): " APP_NAME
if [ -z "$APP_NAME" ]; then
    echo "ERROR: App name cannot be empty"
    exit 1
fi

echo
echo "Creating Heroku app: $APP_NAME"
heroku create $APP_NAME

echo
echo "Adding files to git..."
git add .
git commit -m "Deploy to Heroku"

echo
echo "Deploying to Heroku..."
git push heroku main

if [ $? -eq 0 ]; then
    echo
    echo "========================================"
    echo "    DEPLOYMENT SUCCESSFUL!"
    echo "========================================"
    echo
    echo "Your app is available at:"
    echo "https://$APP_NAME.herokuapp.com"
    echo
    echo "Opening app in browser..."
    heroku open --app $APP_NAME
else
    echo
    echo "ERROR: Deployment failed!"
    echo "Check the logs with: heroku logs --tail --app $APP_NAME"
fi
