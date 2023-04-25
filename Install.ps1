# Set the path to the configuration.xml file
$configFilePath = "configuration.xml"

# Run the command to download the configuration file
Start-Process -FilePath "setup.exe" -ArgumentList "/download", $configFilePath -Wait

# Run the command to configure the application using the downloaded configuration file
Start-Process -FilePath "setup.exe" -ArgumentList "/configure", $configFilePath -Wait
