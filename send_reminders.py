"""
Sends the day's automated overdue-advance reminder emails.

Run manually to test:
    python send_reminders.py

Run automatically via the systemd timer in deploy/ (preferred), or via cron
if you're not using systemd:
    30 8 * * *  cd /opt/advance-register && ./venv/bin/python send_reminders.py >> reminders.log 2>&1

This reuses the exact same eligibility query, cooldown, and email template
as the "Remind all overdue" button in the app (see /api/advances/remind/bulk
in app.py) by calling into app.py directly rather than duplicating the
logic -- the two paths share one cooldown so a person can't be emailed
twice in the same day just because both the button and the timer fired.
"""

import sys
from datetime import datetime, timedelta

from app import app, get_db, build_reminder, send_email, smtp_configured, SETTLEMENT_GRACE_DAYS
import os


def main():
    with app.app_context():
        if not smtp_configured():
            print("SMTP is not configured (SMTP_HOST/SMTP_USER/SMTP_PASS) -- nothing to do.")
            sys.exit(1)

        cooldown_days = int(os.environ.get("REMINDER_COOLDOWN_DAYS", "7"))
        db = get_db()
        cutoff = (datetime.utcnow() - timedelta(days=cooldown_days)).isoformat()
        rows = db.execute(
            """
            SELECT * FROM advances
            WHERE status = 'pending'
              AND email IS NOT NULL AND trim(email) != ''
              AND conf_end_date IS NOT NULL
              AND date(conf_end_date, ?) < date('now')
              AND (last_reminded_at IS NULL OR last_reminded_at < ?)
            """,
            (f"+{SETTLEMENT_GRACE_DAYS} days", cutoff),
        ).fetchall()

        print(f"[{datetime.utcnow().isoformat()}] {len(rows)} advance(s) eligible for a reminder.")

        sent_count, failed_count = 0, 0
        for row in rows:
            subject, body = build_reminder(row)
            ok, err = send_email(row["email"], subject, body)
            if ok:
                db.execute(
                    "UPDATE advances SET last_reminded_at = ? WHERE id = ?",
                    (datetime.utcnow().isoformat(), row["id"]),
                )
                sent_count += 1
                print(f"  sent -> {row['name']} <{row['email']}> ({row['voucher_no']})")
            else:
                failed_count += 1
                print(f"  FAILED -> {row['name']} <{row['email']}> ({row['voucher_no']}): {err}")
        db.commit()

        print(f"Done. {sent_count} sent, {failed_count} failed.")
        if failed_count:
            sys.exit(1)


if __name__ == "__main__":
    main()
