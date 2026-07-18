# Deploying the Advance Register

This app has two parts: a Flask API backend (`app.py`) and a single static
frontend file (`index.html`). In production they're served differently from
how you run them locally:

- Locally: `python app.py` runs Flask's own dev server, and you just open
  `index.html` directly in a browser.
- In production: **gunicorn** runs the backend (Flask's dev server is not
  safe to expose to the network), and **nginx** serves `index.html` and
  proxies `/api/*` requests through to gunicorn, terminating HTTPS.

```
Browser  →  nginx (HTTPS, port 443)  →  gunicorn (127.0.0.1:5000)  →  Flask app  →  SQLite
                │
                └─ serves index.html directly
```

## 1. Server prerequisites

A small Linux VM is plenty for this app's traffic (an internal accounts
tool, not a public site) — 1 vCPU / 1GB RAM is enough. Any of DigitalOcean,
AWS Lightsail, or your institution's own server works. You'll need:

- Python 3.10+
- nginx
- A domain name pointed at the server (for HTTPS — Let's Encrypt needs one;
  a bare IP address can't get a real certificate)

## 2. Get the code onto the server

```bash
sudo useradd -r -m -d /opt/advance-register advanceapp
sudo mkdir -p /opt/advance-register/frontend
# copy app.py, wsgi.py, gunicorn.conf.py, requirements.txt, send_reminders.py
# into /opt/advance-register, and index.html into /opt/advance-register/frontend
sudo chown -R advanceapp:advanceapp /opt/advance-register
```

## 3. Python environment

```bash
sudo -u advanceapp python3 -m venv /opt/advance-register/venv
sudo -u advanceapp /opt/advance-register/venv/bin/pip install -r /opt/advance-register/requirements.txt
```

## 4. Configure `.env`

```bash
sudo -u advanceapp cp /opt/advance-register/_env.example /opt/advance-register/.env
sudo -u advanceapp nano /opt/advance-register/.env
```

Fill in real values, and importantly for production:

- `CORS_ORIGINS=https://advances.yourinstitute.edu` — your real frontend
  domain. Leaving this as `*` in production means any website on the
  internet can call your API from a logged-in admin's browser.
- `FLASK_DEBUG` — leave unset or `0`. This must never be `1` outside your
  own laptop.
- `SMTP_HOST` / `SMTP_USER` / `SMTP_PASS` / `FROM_EMAIL` — needed for both
  the manual "Remind" button and the automated daily reminder job.
- `ADMIN_BOOTSTRAP_PASSWORD` — change this from the example value before
  first run; it's used to create the initial admin account.

## 5. Run the backend as a service (systemd)

```bash
sudo cp /opt/advance-register/deploy/advance-register.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now advance-register
sudo systemctl status advance-register
```

Logs: `sudo journalctl -u advance-register -f`

## 6. Set up nginx + HTTPS

```bash
sudo cp /opt/advance-register/deploy/nginx.conf /etc/nginx/sites-available/advance-register
# edit the file: replace advances.yourinstitute.edu with your real domain
sudo ln -s /etc/nginx/sites-available/advance-register /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx

# Get a free certificate and auto-configure HTTPS:
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d advances.yourinstitute.edu
```

Certbot renews automatically via a systemd timer it installs — no further
action needed.

## 7. Automated reminder emails

```bash
sudo cp /opt/advance-register/deploy/advance-register-reminders.service /etc/systemd/system/
sudo cp /opt/advance-register/deploy/advance-register-reminders.timer /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now advance-register-reminders.timer

# Test it fires correctly without waiting for the schedule:
sudo systemctl start advance-register-reminders.service
sudo journalctl -u advance-register-reminders -f
```

By default this runs daily at 08:30 server time and emails everyone
overdue who hasn't been reminded in the last 7 days (`REMINDER_COOLDOWN_DAYS`
in `.env`). Adjust the schedule in the `.timer` file if needed.

## 8. Firewall

Only 80 and 443 (nginx) need to be open to the internet. Everything else
(gunicorn on 5000, SSH) should not be reachable from outside:

```bash
sudo ufw allow 22/tcp
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw enable
```

## 9. Database backups

The whole database is one file: `/opt/advance-register/advance_register.db`.
Back it up off-server daily — a single file is a single point of failure.
A simple cron entry works fine at this scale:

```bash
# /etc/cron.d/advance-register-backup
0 2 * * * advanceapp cp /opt/advance-register/advance_register.db /opt/advance-register/backups/advance_register_$(date +\%Y\%m\%d).db
```

Prune old backups periodically, and copy them somewhere off the server
(cloud storage, another machine) rather than trusting a second copy on the
same disk.

## Updating the app later

```bash
# copy the new app.py / index.html over
sudo systemctl restart advance-register
```

`init_db()` runs its migrations automatically on every startup, so schema
changes apply themselves — no manual migration step needed.
