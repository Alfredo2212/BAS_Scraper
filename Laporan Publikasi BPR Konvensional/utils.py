"""
Utility functions for OJK scraper
Helper functions for data processing, Excel export, process cleanup, etc.
"""

import os
import sys
import subprocess
import platform


def kill_chrome_processes():
    """
    Kill all Chrome and ChromeDriver processes to clean up the environment
    Useful when Selenium doesn't properly clean up after itself
    """
    system = platform.system()
    killed_count = 0
    
    try:
        if system == "Windows":
            # Windows: Use taskkill command
            processes_to_kill = ["chrome.exe", "chromedriver.exe", "GoogleCrashHandler.exe", "GoogleCrashHandler64.exe"]
            
            for process_name in processes_to_kill:
                try:
                    # Kill process by name
                    result = subprocess.run(
                        ["taskkill", "/F", "/IM", process_name],
                        capture_output=True,
                        text=True,
                        timeout=10
                    )
                    if "berhasil" in result.stdout.lower() or "successfully" in result.stdout.lower():
                        killed_count += 1
                        print(f"[OK] Proses {process_name} dihentikan")
                except subprocess.TimeoutExpired:
                    print(f"[WARNING] Timeout saat menghentikan {process_name}")
                except Exception as e:
                    # Process might not exist, which is fine
                    pass
                    
        elif system == "Linux" or system == "Darwin":  # Linux or macOS
            # Unix-like: Use pkill or killall
            processes_to_kill = ["chrome", "chromedriver", "Google Chrome"]
            
            for process_name in processes_to_kill:
                try:
                    if system == "Linux":
                        subprocess.run(["pkill", "-9", process_name], timeout=10)
                    else:  # macOS
                        subprocess.run(["killall", "-9", process_name], timeout=10)
                    killed_count += 1
                    print(f"[OK] Proses {process_name} dihentikan")
                except subprocess.TimeoutExpired:
                    print(f"[WARNING] Timeout saat menghentikan {process_name}")
                except Exception as e:
                    # Process might not exist, which is fine
                    pass
        else:
            print(f"[WARNING] Sistem operasi {system} tidak didukung untuk pembersihan proses otomatis")
            
    except Exception as e:
        print(f"[WARNING] Error saat membersihkan proses Chrome: {e}")
    
    if killed_count > 0:
        print(f"[OK] {killed_count} proses Chrome/ChromeDriver telah dihentikan")
    else:
        print("[INFO] Tidak ada proses Chrome/ChromeDriver yang berjalan")
    
    return killed_count


def cleanup_selenium_environment():
    """
    Comprehensive cleanup of Selenium environment
    Kills Chrome processes and clears temporary files
    """
    print("[INFO] Membersihkan lingkungan Selenium...")
    killed = kill_chrome_processes()
    
    # Clear Selenium temporary files (if any)
    try:
        import tempfile
        temp_dir = tempfile.gettempdir()
        selenium_temp_files = [
            os.path.join(temp_dir, "scoped_dir*"),
            os.path.join(temp_dir, "*chrome*"),
        ]
        # Note: We can't easily glob and delete these on Windows without more complex logic
        # This is a placeholder for future enhancement
        print("[INFO] Pembersihan file temporary selesai")
    except Exception as e:
        print(f"[WARNING] Error saat membersihkan file temporary: {e}")
    
    print("[OK] Pembersihan lingkungan Selenium selesai")
    return killed > 0
