@echo off
powershell -Command "Invoke-RestMethod -Method 'Post' -ContentType 'application/x-www-form-urlencoded' -OutFile P:\labels.csv -Uri 'https://order.functionalbio.com/admin/label-maker/?auto-download=1' -Body @{ key='LBJyLa9J1JIPzyiRiKyT52fIdnR1Ug8pwb5sLXkgWDzK0QEbOiRThdyOCCJxAw8' }"
