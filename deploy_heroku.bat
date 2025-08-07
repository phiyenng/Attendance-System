@echo off
echo ========================================
echo    DEPLOY ATTENDANCE SYSTEM TO HEROKU
echo ========================================
echo.

echo Checking if Heroku CLI is installed...
heroku --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Heroku CLI not found!
    echo Please install from: https://devcenter.heroku.com/articles/heroku-cli
    pause
    exit /b 1
)

echo Heroku CLI found!
echo.

echo Please login to Heroku:
heroku login
if %errorlevel% neq 0 (
    echo ERROR: Failed to login to Heroku
    pause
    exit /b 1
)

echo.
set /p APP_NAME="Enter your Heroku app name (e.g., my-attendance-system): "
if "%APP_NAME%"=="" (
    echo ERROR: App name cannot be empty
    pause
    exit /b 1
)

echo.
echo Creating Heroku app: %APP_NAME%
heroku create %APP_NAME%

echo.
echo Adding files to git...
git add .
git commit -m "Deploy to Heroku"

echo.
echo Deploying to Heroku...
git push heroku main

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo    DEPLOYMENT SUCCESSFUL!
    echo ========================================
    echo.
    echo Your app is available at:
    echo https://%APP_NAME%.herokuapp.com
    echo.
    echo Opening app in browser...
    heroku open --app %APP_NAME%
) else (
    echo.
    echo ERROR: Deployment failed!
    echo Check the logs with: heroku logs --tail --app %APP_NAME%
)

echo.
pause
