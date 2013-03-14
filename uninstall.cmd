@echo off
REG DELETE HKEY_CURRENT_USER\Software\Vichamp\GTD /f
REG DELETE HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run /v GTD /f
