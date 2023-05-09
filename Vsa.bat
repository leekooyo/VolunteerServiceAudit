@echo off
title Vsa
set "_root=%~dp0"
set "PATH=%_root%Python39;%_root%Python39\Scripts;%PATH%"
cd "%_root%Vsa/Vsa"
python vsa.py