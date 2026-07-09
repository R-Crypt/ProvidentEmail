"""
Tests for the email classifier
"""
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from classifier import EmailClassifier


def test_classifier():
    clf = EmailClassifier()

    test_cases = [
        {
            "subject": "PO #44592 - Cardboard Boxes Order",
            "body": "Please find attached the purchase order for 5000 units. Delivery required by July 15th.",
            "expected": "purchase_order"
        },
        {
            "subject": "Request for Quotation - Custom Packaging",
            "body": "Can you provide pricing for 10,000 units with our logo? We need samples.",
            "expected": "enquiry"
        },
        {
            "subject": "Invoice #INV-0892 - Payment Due",
            "body": "Please find attached invoice for $12,500. Payment terms: Net 30 days.",
            "expected": "invoice"
        },
        {
            "subject": "Shipment Update - Tracking #TRK998877",
            "body": "Your shipment has been dispatched. Tracking number: TRK998877. Estimated delivery: July 2.",
            "expected": "shipping"
        },
        {
            "subject": "Meeting next week",
            "body": "Can we schedule a call to discuss Q3 requirements?",
            "expected": "general"
        }
    ]

    passed = 0
    failed = 0

    for test in test_cases:
        result = clf.classify(test["subject"], test["body"])
        actual = result["category"]
        expected = test["expected"]

        if actual == expected:
            passed += 1
            print(f"✓ PASS: '{test['subject'][:40]}...' -> {actual} ({result['confidence']}%)")
        else:
            failed += 1
            print(f"✗ FAIL: '{test['subject'][:40]}...' -> expected {expected}, got {actual}")

    print(f"\nResults: {passed}/{len(test_cases)} passed")
    return failed == 0


if __name__ == "__main__":
    success = test_classifier()
    sys.exit(0 if success else 1)
