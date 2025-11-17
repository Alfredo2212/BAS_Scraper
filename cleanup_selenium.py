"""
Script untuk membersihkan lingkungan Selenium
Menutup semua proses Chrome dan ChromeDriver yang masih berjalan
"""

import sys
from pathlib import Path

# Add the module directory to path
sys.path.insert(0, str(Path(__file__).parent))

try:
    from Laporan_Publikasi_BPR_Konvensional.utils import cleanup_selenium_environment, kill_chrome_processes
except ImportError:
    # Try alternative import
    try:
        import importlib.util
        module_path = Path(__file__).parent / "Laporan Publikasi BPR Konvensional" / "utils.py"
        spec = importlib.util.spec_from_file_location("utils_module", module_path)
        utils_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(utils_module)
        cleanup_selenium_environment = utils_module.cleanup_selenium_environment
        kill_chrome_processes = utils_module.kill_chrome_processes
    except Exception as e:
        print(f"[ERROR] Tidak dapat mengimpor modul utils: {e}")
        sys.exit(1)


def main():
    """Main function untuk membersihkan lingkungan Selenium"""
    print("=" * 60)
    print("Pembersihan Lingkungan Selenium")
    print("=" * 60)
    print()
    
    print("[INFO] Menghentikan semua proses Chrome dan ChromeDriver...")
    success = cleanup_selenium_environment()
    
    if success:
        print("\n[OK] Pembersihan selesai! Lingkungan telah dibersihkan.")
    else:
        print("\n[INFO] Tidak ada proses yang perlu dihentikan.")
    
    print("\n[INFO] Tekan ENTER untuk keluar...")
    try:
        input()
    except EOFError:
        pass


if __name__ == "__main__":
    main()

