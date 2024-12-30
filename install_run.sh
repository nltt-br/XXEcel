#!/bin/bash

if [ ${EUID:-0} -ne 0 ] || [ "$(id -u)" -ne 0 ] || [groups $(whoami) | grep -q '\bdocker\b']; then
    echo "You must be root to do this or group docker" 1>&2
    exit 100
else
    docker build -t xxecel .
    docker run -p 5000:5000 xxecel
fi