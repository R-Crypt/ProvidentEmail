"""
Provident Packaging Email Classifier - Main Automation Scheduler
Runs every 15 minutes to check, classify, and organize emails
"""
import time
import logging
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger
from config import config
from graph_client import GraphClient
from classifier import EmailClassifier
from database import Database

# Set up logging
logging.basicConfig(
    level=getattr(logging, config.LOG_LEVEL),
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class EmailAutomation:
    """Main automation engine that orchestrates email classification"""

    def __init__(self):
        self.graph = GraphClient()
        self.classifier = EmailClassifier()
        self.db = Database()
        self.stats = {
            "processed": 0,
            "errors": 0,
            "start_time": datetime.utcnow()
        }

    def run_once(self):
        """
        Process one batch of emails across all users
        This is the main workhorse function
        """
        logger.info("=" * 60)
        logger.info("Starting email classification check...")

        batch_stats = {
            "processed": 0,
            "failed": 0,
            "by_category": {}
        }

        try:
            # Get all users in the organization
            users = self.graph.get_users()
            logger.info(f"Found {len(users)} users to process")

            # Check emails from last 30 minutes
            since = datetime.utcnow() - timedelta(minutes=config.LOOKBACK_MINUTES)

            for user in users:
                user_email = user.get("mail") or user.get("userPrincipalName")
                user_id = user["id"]

                logger.info(f"Checking emails for: {user_email}")

                # Get unread emails
                emails = self.graph.get_unread_emails(user_id, since)

                if not emails:
                    logger.info(f"  No new emails for {user_email}")
                    continue

                logger.info(f"  Found {len(emails)} unread emails")

                # Ensure Outlook categories exist for this user
                self._ensure_categories(user_id)

                for email in emails:
                    message_id = email["id"]

                    # Skip if already processed
                    if self.db.is_processed(message_id):
                        logger.debug(f"  Skipping already processed: {message_id}")
                        continue

                    try:
                        # Extract email data
                        subject = email.get("subject", "(No subject)")
                        body = email.get("body", {}).get("content", "")
                        body_preview = email.get("bodyPreview", "")[:200]
                        sender = email.get("from", {}).get("emailAddress", {}).get("address", "unknown")
                        received_at = self._parse_datetime(email.get("receivedDateTime"))

                        # Classify the email
                        result = self.classifier.classify(subject, body)
                        category = result["category"]
                        confidence = result["confidence"]
                        reason = result["reason"]

                        # Get Outlook category name
                        outlook_cat = config.CATEGORIES[category]["outlook_category"]

                        # Apply category in Outlook
                        category_applied = self.graph.apply_category(
                            user_id, message_id, outlook_cat
                        )

                        # Record in database
                        self.db.record_classification(
                            message_id=message_id,
                            user_email=user_email,
                            sender=sender,
                            subject=subject,
                            body_preview=body_preview,
                            category=category,
                            confidence=confidence,
                            reason=reason,
                            outlook_category=outlook_cat if category_applied else None,
                            received_at=received_at,
                            extracted_data=result.get("extracted_data"),
                            response_draft=result.get("response_draft"),
                            source_folder=email.get("source_folder", "Inbox")
                        )

                        # Auto-Reply Check
                        if getattr(config, "AUTO_REPLY_ENABLED", False):
                            logger.info(f"  [Auto-Reply Active] Simulating auto-reply for: {message_id}")
                            self.db.record_reply_sent(message_id, result.get("response_draft") or "Thank you for your message. We have received it and are processing it.")

                        # Update stats
                        batch_stats["processed"] += 1
                        batch_stats["by_category"][category] = batch_stats["by_category"].get(category, 0) + 1

                        logger.info(
                            f"  ✓ Classified: '{subject[:50]}...' -> "
                            f"{config.CATEGORIES[category]['display_name']} "
                            f"({confidence}%)"
                        )

                    except Exception as e:
                        batch_stats["failed"] += 1
                        logger.error(f"  ✗ Failed to process email {message_id}: {e}")
                        continue

            # Update overall stats
            self.stats["processed"] += batch_stats["processed"]

            logger.info("=" * 60)
            logger.info(f"Batch complete: {batch_stats['processed']} processed, "
                       f"{batch_stats['failed']} failed")
            if batch_stats["by_category"]:
                logger.info(f"By category: {batch_stats['by_category']}")
            logger.info("=" * 60)

        except Exception as e:
            self.stats["errors"] += 1
            logger.error(f"Critical error in run_once: {e}", exc_info=True)

    def _ensure_categories(self, user_id: str):
        """Ensure all required Outlook categories exist for a user"""
        for cat_id, cat_info in config.CATEGORIES.items():
            if cat_id == "general":
                continue
            try:
                self.graph.create_outlook_category(
                    user_id, 
                    cat_info["outlook_category"], 
                    cat_info["color"]
                )
            except Exception as e:
                logger.warning(f"Could not create category {cat_info['outlook_category']}: {e}")

    def _parse_datetime(self, dt_str: str) -> datetime:
        """Parse ISO datetime string"""
        if not dt_str:
            return datetime.utcnow()
        try:
            # Handle ISO format with Z suffix
            dt_str = dt_str.replace("Z", "+00:00")
            return datetime.fromisoformat(dt_str)
        except:
            return datetime.utcnow()

    def print_stats(self):
        """Print overall statistics"""
        runtime = datetime.utcnow() - self.stats["start_time"]
        logger.info("\n" + "=" * 60)
        logger.info("OVERALL STATISTICS")
        logger.info("=" * 60)
        logger.info(f"Runtime: {runtime}")
        logger.info(f"Total processed: {self.stats['processed']}")
        logger.info(f"Total errors: {self.stats['errors']}")
        logger.info("=" * 60)

    def start(self):
        """Start the background scheduler"""
        logger.info("=" * 60)
        logger.info("PROVIDENT PACKAGING EMAIL CLASSIFIER")
        logger.info("=" * 60)
        logger.info(f"Check interval: {config.CHECK_INTERVAL_MINUTES} minutes")
        logger.info(f"Lookback: {config.LOOKBACK_MINUTES} minutes")
        logger.info(f"AI Model: {config.OPENAI_MODEL}")
        logger.info("=" * 60)

        # Validate configuration
        if not config.validate():
            logger.error("Configuration validation failed. Exiting.")
            return

        # Run immediately on startup
        self.run_once()

        # Set up scheduler
        scheduler = BackgroundScheduler()
        scheduler.add_job(
            self.run_once,
            trigger=IntervalTrigger(minutes=config.CHECK_INTERVAL_MINUTES),
            id="email_classifier",
            name="Email Classification Job",
            replace_existing=True
        )

        scheduler.start()
        logger.info(f"Scheduler started. Running every {config.CHECK_INTERVAL_MINUTES} minutes.")
        logger.info("Press Ctrl+C to stop.\n")

        try:
            # Keep the main thread alive
            while True:
                time.sleep(60)
        except KeyboardInterrupt:
            logger.info("\nShutting down...")
            scheduler.shutdown()
            self.print_stats()


if __name__ == "__main__":
    automation = EmailAutomation()
    automation.start()
