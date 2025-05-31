#!/bin/bash
gunicorn --access-logfile=- -w 4 --bind 0.0.0.0:8080 --timeout 1200 'extrator_dgt:app'
