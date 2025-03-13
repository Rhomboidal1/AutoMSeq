@echo off
powershell -Command "Invoke-RestMethod -Method 'Post' -ContentType 'application/x-www-form-urlencoded' -OutFile P:\order_key.txt -Uri 'https://order.functionalbio.com/admin/generate-data-sorting-key/?auto-download=1' -Body @{ key='5Y0Ho2yZs5S6mTjohw8FE6OLvzg4aB7CThusxxWG77G5cyF6kEOY4m4SmxDECzO' }"
