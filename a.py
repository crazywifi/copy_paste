# Core functionality
import os
import re
import time
import json
import sys
import requests
import concurrent.futures
from datetime import datetime
from urllib.parse import urlparse, parse_qs

# GUI components
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, Menu

# Web automation
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# System interactions
import subprocess
import platform
import webbrowser

# Document parsing (for search functionality)
import fitz  # PyMuPDF for PDFs
import docx
import pptx
import pandas as pd

# Optional: Terminal colors (you could remove if not needed)
#import colorama
#from colorama import Fore, Style
#colorama.init()
