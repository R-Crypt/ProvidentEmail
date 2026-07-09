"""
Provident Packaging Email Classifier - Integrated Outlook Mockup Dashboard
Highly customized for http://providentpackaging.com/
- Professional Fluent UI styles (flat designs, high contrast blue elements, clean grids)
- Complete removal of emojis
- Clean inline vector SVG icons
- Dynamic drag-to-resize sidebar split
- Action Triage list (with category filters)
- Product Catalog, cross-sell integrations, and PO/RFQ calculators
"""
import os
import sys
import json
from flask import Flask, render_template_string, jsonify, request, send_from_directory

# Add src to path if needed
sys.path.insert(0, os.path.dirname(__file__))

from database import Database
from config import config
from classifier import EmailClassifier

app = Flask(__name__)
db = Database()

DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mail - Outlook Provident Workspace</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Segoe+UI:wght@300;400;600;700&family=Outfit:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@500;600;700&display=swap" rel="stylesheet">
    
    <style>
        :root {
            /* Outlook Color Palette */
            --outlook-blue: #0f6cbd;
            --outlook-blue-hover: #1178cc;
            --outlook-dark-blue: #005a9e;
            --gray-border: #edebe9;
            --gray-border-dark: #d2d0ce;
            --bg-light: #f3f2f1;
            --bg-white: #ffffff;
            
            --text-primary: #323130;
            --text-secondary: #605e5c;
            --text-muted: #a19f9d;
            --hover-bg: #f3f2f1;
            --selected-bg: #edebe9;
            
            /* Category Colors */
            --cat-po: #107c41;
            --cat-po-bg: #dff6dd;
            --cat-enquiry: #0078d4;
            --cat-enquiry-bg: #deecf9;
            --cat-invoice: #a80000;
            --cat-invoice-bg: #fde7e9;
            --cat-shipping: #8764b8;
            --cat-shipping-bg: #ebd3f8;
            --cat-general: #605e5c;
            --cat-general-bg: #f3f2f1;
            
            /* Provident Accent */
            --provident-gold: #b39257;
            --provident-dark: #1b2a47;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            background-color: var(--bg-white);
            color: var(--text-primary);
            height: 100vh;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }

        /* --- 1. OUTLOOK BLUE HEADER --- */
        .outlook-header {
            height: 48px;
            background-color: var(--outlook-blue);
            color: white;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 16px;
            font-size: 0.9rem;
            z-index: 100;
            flex-shrink: 0;
        }

        .header-left {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .waffle-menu {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 3px;
            width: 16px;
            height: 16px;
            cursor: pointer;
        }

        .waffle-dot {
            background-color: white;
            border-radius: 50%;
            width: 3.5px;
            height: 3.5px;
        }

        .app-title {
            font-weight: 600;
            font-size: 1.05rem;
            letter-spacing: 0.2px;
        }

        .header-search {
            flex-grow: 1;
            max-width: 460px;
            margin: 0 20px;
            position: relative;
        }

        .search-input {
            width: 100%;
            height: 32px;
            border-radius: 4px;
            border: none;
            background-color: rgba(255, 255, 255, 0.15);
            padding: 0 12px 0 36px;
            color: white;
            font-size: 0.85rem;
            outline: none;
            transition: background-color 0.15s;
        }

        .search-input::placeholder {
            color: rgba(255, 255, 255, 0.8);
        }

        .search-input:focus {
            background-color: white;
            color: var(--text-primary);
        }

        .header-right {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .header-right-icon {
            cursor: pointer;
            width: 32px;
            height: 32px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 4px;
            transition: background-color 0.2s;
            color: white;
        }

        .header-right-icon:hover {
            background-color: rgba(255, 255, 255, 0.1);
        }

        .avatar-circle {
            background-color: #8b5cf6;
            color: white;
            font-size: 0.8rem;
            font-weight: 600;
            width: 30px !important;
            height: 30px !important;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        /* --- 2. OUTLOOK SUB-HEADER / RIBBON --- */
        .outlook-ribbon {
            height: 48px;
            border-bottom: 1px solid var(--gray-border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 16px;
            background-color: var(--bg-white);
            flex-shrink: 0;
        }

        .ribbon-left {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .new-email-btn {
            background-color: var(--outlook-blue);
            color: white;
            border: none;
            border-radius: 4px;
            padding: 6px 14px;
            font-size: 0.85rem;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: background-color 0.15s;
        }

        .new-email-btn:hover {
            background-color: var(--outlook-blue-hover);
        }

        .ribbon-divider {
            width: 1px;
            height: 20px;
            background-color: var(--gray-border-dark);
            margin: 0 4px;
        }

        .ribbon-action-btn {
            background: none;
            border: none;
            color: var(--text-primary);
            padding: 6px 12px;
            font-size: 0.85rem;
            cursor: pointer;
            border-radius: 4px;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: background-color 0.15s;
        }

        .ribbon-action-btn:hover {
            background-color: var(--hover-bg);
        }

        .ribbon-action-btn.provident-toggle-btn {
            background-color: rgba(15, 108, 189, 0.08);
            border: 1px solid rgba(15, 108, 189, 0.2);
            color: var(--outlook-blue);
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .ribbon-action-btn.provident-toggle-btn:hover {
            background-color: rgba(15, 108, 189, 0.15);
        }

        /* --- 3. MAIN CONTENT GRID --- */
        .outlook-body {
            flex-grow: 1;
            display: flex;
            overflow: hidden;
            background-color: var(--bg-white);
            position: relative;
        }

        /* leftmost sidebar */
        .outlook-left-rail {
            width: 48px;
            border-right: 1px solid var(--gray-border);
            background-color: var(--bg-white);
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 12px 0;
            gap: 20px;
            color: var(--text-secondary);
            flex-shrink: 0;
        }

        .left-rail-icon {
            cursor: pointer;
            width: 32px;
            height: 32px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 4px;
            transition: background-color 0.15s;
        }

        .left-rail-icon:hover {
            background-color: var(--hover-bg);
            color: var(--outlook-blue);
        }

        .left-rail-icon.active {
            color: var(--outlook-blue);
            background-color: rgba(15, 108, 189, 0.08);
            border-left: 2px solid var(--outlook-blue);
            border-radius: 0 4px 4px 0;
        }

        /* --- 4. INBOX LIST COLUMN --- */
        .inbox-list-column {
            width: 320px;
            border-right: 1px solid var(--gray-border);
            display: flex;
            flex-direction: column;
            overflow: hidden;
            flex-shrink: 0;
        }

        .inbox-tabs {
            display: flex;
            height: 36px;
            border-bottom: 1px solid var(--gray-border);
            padding: 0 16px;
            align-items: center;
            gap: 16px;
            font-size: 0.85rem;
            font-weight: 600;
        }

        .inbox-tab {
            cursor: pointer;
            padding: 8px 0;
            position: relative;
            color: var(--text-secondary);
        }

        .inbox-tab.active {
            color: var(--text-primary);
        }

        .inbox-tab.active::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 2px;
            background-color: var(--outlook-blue);
        }

        .inbox-scroll {
            flex-grow: 1;
            overflow-y: auto;
        }

        .inbox-section-header {
            padding: 10px 16px 4px 16px;
            font-size: 0.75rem;
            font-weight: 700;
            color: var(--outlook-blue);
            text-transform: uppercase;
            letter-spacing: 0.3px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .clear-filter-btn {
            font-size: 0.68rem;
            color: var(--text-muted);
            cursor: pointer;
            text-transform: none;
            font-weight: normal;
        }

        .clear-filter-btn:hover {
            color: var(--outlook-blue);
            text-decoration: underline;
        }

        /* Individual Email Card */
        .email-card {
            display: flex;
            padding: 12px 16px;
            border-bottom: 1px solid rgba(0, 0, 0, 0.03);
            cursor: pointer;
            position: relative;
            transition: background-color 0.15s;
        }

        .email-card:hover {
            background-color: var(--hover-bg);
        }

        .email-card.selected {
            background-color: var(--selected-bg);
        }

        .email-card::before {
            content: '';
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 3px;
            background-color: transparent;
        }

        .email-card.unread::before {
            background-color: var(--outlook-blue);
        }

        .email-avatar {
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background-color: #a19f9d;
            color: white;
            font-weight: 600;
            font-size: 0.85rem;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 12px;
            flex-shrink: 0;
        }

        .email-card-content {
            flex-grow: 1;
            min-width: 0;
        }

        .email-card-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 3px;
        }

        .email-card-sender {
            font-weight: 600;
            font-size: 0.88rem;
            color: var(--text-primary);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .email-card-time {
            font-size: 0.75rem;
            color: var(--text-secondary);
        }

        .email-card-subject {
            font-size: 0.82rem;
            color: var(--text-primary);
            margin-bottom: 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .email-card-body-snippet {
            font-size: 0.8rem;
            color: var(--text-secondary);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .email-card-meta {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 6px;
        }

        .outlook-badge {
            font-size: 0.72rem;
            font-weight: 600;
            padding: 2px 8px;
            border-radius: 3px;
            text-transform: uppercase;
        }

        .outlook-badge.purchase_order { background-color: var(--cat-po-bg); color: var(--cat-po); }
        .outlook-badge.enquiry { background-color: var(--cat-enquiry-bg); color: var(--cat-enquiry); }
        .outlook-badge.invoice { background-color: var(--cat-invoice-bg); color: var(--cat-invoice); }
        .outlook-badge.shipping { background-color: var(--cat-shipping-bg); color: var(--cat-shipping); }
        .outlook-badge.general { background-color: var(--cat-general-bg); color: var(--cat-general); }

        .attachment-icon-svg {
            color: var(--text-secondary);
            display: flex;
            align-items: center;
        }

        /* --- 5. CENTRAL READING PANE --- */
        .outlook-reading-pane {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            min-width: 200px;
        }

        .empty-reading-pane {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            color: var(--text-secondary);
            text-align: center;
            padding: 2rem;
        }

        .empty-pane-icon-svg {
            color: var(--text-muted);
            margin-bottom: 1.5rem;
            opacity: 0.35;
        }

        .empty-pane-title {
            font-size: 1.1rem;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 4px;
        }

        .empty-pane-subtitle {
            font-size: 0.85rem;
            color: var(--text-muted);
        }

        .email-details-view {
            display: flex;
            flex-direction: column;
            height: 100%;
            overflow-y: auto;
            padding: 24px;
        }

        .email-view-header {
            margin-bottom: 20px;
            padding-bottom: 16px;
            border-bottom: 1px solid var(--gray-border);
        }

        .email-view-subject {
            font-family: 'Segoe UI', sans-serif;
            font-size: 1.4rem;
            font-weight: 400;
            color: var(--text-primary);
            margin-bottom: 14px;
        }

        .email-view-meta-row {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .email-view-sender-avatar {
            width: 40px;
            height: 40px;
            border-radius: 50%;
            background-color: var(--outlook-blue);
            color: white;
            font-weight: 600;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.95rem;
        }

        .email-view-sender-info {
            flex-grow: 1;
            line-height: 1.4;
        }

        .email-view-sender-name {
            font-weight: 600;
            font-size: 0.9rem;
        }

        .email-view-sender-address {
            font-size: 0.8rem;
            color: var(--text-secondary);
        }

        .email-view-date {
            font-size: 0.8rem;
            color: var(--text-secondary);
        }

        .email-view-category-row {
            margin-top: 10px;
            display: flex;
            gap: 8px;
        }

        .email-view-body {
            font-size: 0.92rem;
            line-height: 1.6;
            color: var(--text-primary);
            white-space: pre-wrap;
            margin-bottom: 24px;
            font-family: Arial, Helvetica, sans-serif;
        }

        /* --- 6. RESIZABLE DRAG HANDLE --- */
        .resize-handle {
            width: 5px;
            background-color: var(--gray-border);
            cursor: col-resize;
            transition: background-color 0.2s;
            z-index: 150;
            flex-shrink: 0;
            position: relative;
        }

        .resize-handle:hover, .resize-handle.active {
            background-color: var(--outlook-blue);
            width: 7px;
        }

        /* --- 7. INTEGRATED WORKSPACE SIDEPANEL --- */
        .outlook-side-panel {
            width: 380px;
            background-color: var(--bg-white);
            display: flex;
            flex-direction: column;
            overflow: hidden;
            flex-shrink: 0;
            min-width: 280px;
            max-width: 600px;
        }

        .outlook-side-panel.collapsed {
            width: 0 !important;
            min-width: 0 !important;
            display: none !important;
        }

        .panel-header {
            height: 48px;
            padding: 0 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 1px solid var(--gray-border);
            background-color: var(--bg-white);
            flex-shrink: 0;
        }

        .panel-logo {
            font-size: 0.9rem;
            font-weight: 700;
            color: white;
            background-color: var(--provident-dark);
            padding: 3px 8px;
            border-radius: 2px;
            border-left: 3px solid var(--provident-gold);
            font-family: 'Outfit', sans-serif;
            letter-spacing: 0.5px;
        }

        .close-panel-btn {
            background: none;
            border: none;
            font-size: 1rem;
            color: var(--text-secondary);
            cursor: pointer;
            width: 28px;
            height: 28px;
            border-radius: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .close-panel-btn:hover {
            background-color: var(--hover-bg);
            color: var(--text-primary);
        }

        .panel-scroll {
            flex-grow: 1;
            overflow-y: auto;
            padding: 16px;
            display: flex;
            flex-direction: column;
            gap: 20px;
            background-color: var(--bg-white);
        }

        /* Triage Checklist Widget */
        .triage-widget {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .triage-description {
            font-size: 0.82rem;
            color: var(--text-secondary);
            line-height: 1.4;
            margin-bottom: 4px;
        }

        .triage-checklist-item {
            display: flex;
            align-items: center;
            padding: 12px 14px;
            background-color: var(--bg-light);
            border: 1px solid var(--gray-border);
            border-radius: 4px;
            cursor: pointer;
            transition: all 0.2s;
            position: relative;
        }

        .triage-checklist-item:hover {
            border-color: var(--outlook-blue);
            background-color: var(--bg-white);
            transform: translateY(-1px);
        }

        .triage-checklist-item.active-filter {
            border-color: var(--outlook-blue);
            background-color: rgba(15, 108, 189, 0.05);
            box-shadow: inset 0 0 0 1px var(--outlook-blue);
        }

        .triage-checklist-item.active-filter::after {
            content: 'Filter Active';
            position: absolute;
            right: 14px;
            top: 6px;
            font-size: 0.62rem;
            color: var(--outlook-blue);
            font-weight: 700;
            text-transform: uppercase;
        }

        .triage-item-bullet {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            margin-right: 14px;
            flex-shrink: 0;
        }

        .triage-item-info {
            flex-grow: 1;
            line-height: 1.35;
        }

        .triage-item-title {
            font-size: 0.85rem;
            font-weight: 600;
            color: var(--text-primary);
        }

        .triage-item-sub {
            font-size: 0.76rem;
            color: var(--text-secondary);
        }

        .triage-item-count {
            background-color: var(--gray-border-dark);
            color: var(--text-primary);
            font-size: 0.78rem;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 12px;
            flex-shrink: 0;
            margin-left: 8px;
        }

        .triage-checklist-item.active-filter .triage-item-count {
            background-color: var(--outlook-blue);
            color: white;
        }

        /* --- ACTIVE COPILOT ASSISTANT VIEW --- */
        .copilot-assistant-view {
            display: flex;
            flex-direction: column;
            gap: 14px;
        }

        /* Tab Layout */
        .copilot-tabs {
            display: flex;
            border-bottom: 1px solid var(--gray-border);
            margin-bottom: 8px;
            gap: 16px;
        }

        .copilot-tab {
            cursor: pointer;
            padding: 6px 0 8px 0;
            font-size: 0.82rem;
            font-weight: 600;
            color: var(--text-secondary);
            position: relative;
        }

        .copilot-tab.active {
            color: var(--outlook-blue);
        }

        .copilot-tab.active::after {
            content: '';
            position: absolute;
            bottom: -1px;
            left: 0;
            right: 0;
            height: 2px;
            background-color: var(--outlook-blue);
        }

        .copilot-card {
            border: 1px solid var(--gray-border);
            border-radius: 4px;
            padding: 14px;
            background-color: var(--bg-white);
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            display: flex;
            flex-direction: column;
            gap: 10px;
        }

        .copilot-section-title {
            font-size: 0.8rem;
            font-weight: 700;
            text-transform: uppercase;
            color: var(--provident-dark);
            border-left: 3px solid var(--provident-gold);
            padding-left: 8px;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        /* Re-classify dropdown */
        .reclassify-select {
            width: 100%;
            height: 30px;
            padding: 0 8px;
            border-radius: 4px;
            border: 1px solid var(--gray-border-dark);
            font-size: 0.82rem;
            outline: none;
            background-color: var(--bg-white);
            color: var(--text-primary);
            cursor: pointer;
        }

        /* Extracted metadata fields */
        .extracted-field-row {
            display: flex;
            flex-direction: column;
            gap: 4px;
            padding-bottom: 8px;
            border-bottom: 1px solid rgba(0, 0, 0, 0.02);
        }

        .extracted-field-row:last-child {
            border-bottom: none;
            padding-bottom: 0;
        }

        .field-label {
            font-size: 0.72rem;
            color: var(--text-secondary);
            font-weight: 600;
            text-transform: uppercase;
        }

        .field-value-container {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 8px;
        }

        .field-value {
            font-size: 0.84rem;
            font-weight: 600;
            color: var(--text-primary);
        }

        .copy-btn {
            background: none;
            border: none;
            color: var(--outlook-blue);
            cursor: pointer;
            font-size: 0.85rem;
            padding: 3px 6px;
            border-radius: 4px;
            transition: all 0.15s;
            display: flex;
            align-items: center;
        }

        .copy-btn:hover {
            background-color: var(--hover-bg);
        }

        .copy-btn.copied {
            color: var(--cat-po);
        }

        /* ERP Checklist Styling */
        .erp-checklist {
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin-top: 4px;
        }

        .erp-check-item {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.82rem;
            color: var(--text-primary);
        }

        .erp-check-item input[type="checkbox"] {
            width: 14px;
            height: 14px;
            cursor: pointer;
        }

        .erp-push-btn {
            background-color: var(--cat-po);
            color: white;
            border: none;
            border-radius: 4px;
            padding: 8px 12px;
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
            text-align: center;
            transition: background-color 0.15s;
            margin-top: 6px;
        }

        .erp-push-btn:hover {
            background-color: #0b5c30;
        }

        .erp-push-btn.pushed {
            background-color: var(--outlook-blue);
        }

        /* Dynamic Quotation Pricing Calculator */
        .calc-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            margin-top: 4px;
        }

        .calc-input-group {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .calc-input-group label {
            font-size: 0.72rem;
            color: var(--text-secondary);
            font-weight: 600;
        }

        .calc-input-group input, .calc-input-group select {
            height: 28px;
            border: 1px solid var(--gray-border-dark);
            border-radius: 4px;
            padding: 0 6px;
            font-size: 0.82rem;
            outline: none;
        }

        .calc-result-box {
            grid-column: span 2;
            background-color: var(--bg-light);
            border: 1px solid var(--gray-border);
            border-radius: 4px;
            padding: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 4px;
        }

        .calc-total-label {
            font-size: 0.76rem;
            font-weight: 600;
            color: var(--text-secondary);
        }

        .calc-total-val {
            font-size: 1.15rem;
            font-weight: 700;
            color: var(--provident-dark);
        }

        .calc-apply-btn {
            grid-column: span 2;
            background-color: var(--outlook-blue);
            color: white;
            border: none;
            border-radius: 4px;
            padding: 6px 12px;
            font-size: 0.8rem;
            font-weight: 600;
            cursor: pointer;
            transition: background-color 0.15s;
        }

        .calc-apply-btn:hover {
            background-color: var(--outlook-blue-hover);
        }

        /* Product Match & Cross-Sell Widget */
        .catalog-match-box {
            background-color: rgba(179, 146, 87, 0.08);
            border: 1px solid rgba(179, 146, 87, 0.25);
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 4px;
        }

        .match-badge {
            background-color: var(--provident-dark);
            color: white;
            font-size: 0.68rem;
            font-weight: 600;
            padding: 2px 6px;
            border-radius: 2px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            display: inline-block;
            margin-bottom: 6px;
        }

        .cross-sell-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 8px 10px;
            border: 1px solid var(--gray-border);
            border-radius: 4px;
            background-color: var(--bg-white);
            margin-top: 6px;
            transition: all 0.2s;
        }

        .cross-sell-item:hover {
            border-color: var(--provident-gold);
        }

        .cross-sell-info {
            line-height: 1.3;
        }

        .cross-sell-title {
            font-size: 0.82rem;
            font-weight: 600;
            color: var(--text-primary);
        }

        .cross-sell-vendor {
            font-size: 0.72rem;
            color: var(--text-muted);
        }

        .add-cross-btn {
            background: none;
            border: 1px solid var(--outlook-blue);
            color: var(--outlook-blue);
            border-radius: 4px;
            width: 24px;
            height: 24px;
            font-size: 1rem;
            font-weight: 400;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            line-height: 1;
        }

        .add-cross-btn:hover {
            background-color: rgba(15, 108, 189, 0.05);
        }

        .add-cross-btn.added {
            background-color: var(--cat-po);
            border-color: var(--cat-po);
            color: white;
        }

        /* Auto draft text area */
        .draft-textarea {
            width: 100%;
            height: 160px;
            padding: 10px;
            border: 1px solid var(--gray-border-dark);
            border-radius: 4px;
            font-size: 0.82rem;
            line-height: 1.45;
            color: var(--text-primary);
            resize: none;
            outline: none;
            font-family: Arial, Helvetica, sans-serif;
            background-color: #faf9f8;
        }

        .draft-textarea:focus {
            border-color: var(--outlook-blue);
            background-color: var(--bg-white);
        }

        .copy-draft-btn {
            background-color: var(--outlook-blue);
            color: white;
            border: none;
            border-radius: 4px;
            padding: 8px 16px;
            font-size: 0.85rem;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
            text-align: center;
            transition: background-color 0.15s;
            margin-top: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
        }

        .copy-draft-btn:hover {
            background-color: var(--outlook-blue-hover);
        }

        .copy-draft-btn.copied {
            background-color: var(--cat-po);
        }

        /* Back to triage button */
        .back-triage-btn {
            background: none;
            border: 1px solid var(--outlook-blue);
            color: var(--outlook-blue);
            border-radius: 4px;
            padding: 6px 12px;
            font-size: 0.8rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s;
            align-self: flex-start;
        }

        .back-triage-btn:hover {
            background-color: rgba(15, 108, 189, 0.05);
        }

        .triage-quick-stats {
            display: flex;
            justify-content: space-between;
            font-size: 0.78rem;
            color: var(--text-secondary);
            padding: 8px 12px;
            background-color: var(--hover-bg);
            border-radius: 4px;
            font-weight: 600;
        }
    </style>
</head>
<body>

    <!-- 1. HEADER -->
    <header class="outlook-header">
        <div class="header-left">
            <div class="waffle-menu">
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
                <div class="waffle-dot"></div>
            </div>
            <div class="app-title">Outlook</div>
        </div>
        <div class="header-search">
            <svg class="search-icon-svg" viewBox="0 0 20 20" width="16" height="16" fill="white" style="position: absolute; left: 12px; top: 50%; transform: translateY(-50%); pointer-events: none; opacity: 0.85;">
                <path d="M8.5 3a5.5 5.5 0 1 0 5.5 5.5A5.5 5.5 0 0 0 8.5 3zM2 8.5a6.5 6.5 0 1 1 11.4 4.19l4.5 4.5a.5.5 0 0 1-.7.7l-4.5-4.5A6.5 6.5 0 0 1 2 8.5z"/>
            </svg>
            <input type="text" class="search-input" placeholder="Search emails, contacts, or folders">
        </div>
        <div class="header-right">
            <div class="header-right-icon" title="Calendar">
                <svg viewBox="0 0 20 20" width="16" height="16" fill="currentColor">
                    <path d="M5.5 2a.5.5 0 0 1 .5.5V3h8v-.5a.5.5 0 0 1 1 0V3h1.5A1.5 1.5 0 0 1 18 4.5v11a1.5 1.5 0 0 1-1.5 1.5H3.5A1.5 1.5 0 0 1 2 15.5v-11A1.5 1.5 0 0 1 3.5 3H5v-.5A.5.5 0 0 1 5.5 2zM3 7v8.5a.5.5 0 0 0 .5.5h13a.5.5 0 0 0 .5-.5V7H3zm13.5-3H3.5a.5.5 0 0 0-.5.5V6h14v-1.5a.5.5 0 0 0-.5-.5z"/>
                </svg>
            </div>
            <div class="header-right-icon" title="Notifications">
                <svg viewBox="0 0 20 20" width="16" height="16" fill="currentColor">
                    <path d="M10 2a3.02 3.02 0 0 0-3 3v2.85a4.48 4.48 0 0 1-1.07 2.9L5.3 11.5A1 1 0 0 0 6 13h8a1 1 0 0 0 .7-1.7l-.63-.75A4.48 4.48 0 0 1 13 7.85V5a3.02 3.02 0 0 0-3-3zm-1.5 12a1.5 1.5 0 0 0 3 0h-3z"/>
                </svg>
            </div>
            <div class="header-right-icon" title="Settings">
                <svg viewBox="0 0 20 20" width="16" height="16" fill="currentColor">
                    <path d="M9.4 2.04a1 1 0 0 1 1.2 0l.74.55a1.2 1.2 0 0 0 1.26.16l.86-.38a1 1 0 0 1 1.15.26l.66.72a1 1 0 0 1 .15 1.17l-.39.86a1.2 1.2 0 0 0 .16 1.26l.55.74a1 1 0 0 1 0 1.2l-.55.74a1.2 1.2 0 0 0-.16 1.26l.39.86a1 1 0 0 1-.15 1.17l-.66.72a1 1 0 0 1-1.15.26l-.86-.38a1.2 1.2 0 0 0-1.26.16l-.74.55a1 1 0 0 1-1.2 0l-.74-.55a1.2 1.2 0 0 0-1.26-.16l-.86.38a1 1 0 0 1-1.15-.26l-.66-.72a1 1 0 0 1-.15-1.17l.39-.86a1.2 1.2 0 0 0-.16-1.26l-.55-.74a1 1 0 0 1 0-1.2l.55-.74a1.2 1.2 0 0 0 .16-1.26l-.39-.86a1 1 0 0 1 .15-1.17l.66-.72a1 1 0 0 1 1.15-.26l.86.38c.4.18.88.12 1.26-.16l.74-.55zM10 7a3 3 0 1 0 0 6 3 3 0 0 0 0-6z"/>
                </svg>
            </div>
            <div class="avatar-circle">PP</div>
        </div>
    </header>

    <!-- 2. RIBBON COMMAND BAR -->
    <div class="outlook-ribbon">
        <div class="ribbon-left">
            <button class="new-email-btn">
                <svg viewBox="0 0 20 20" width="12" height="12" fill="white" style="margin-right: 4px;">
                    <path d="M10 2a.75.75 0 0 1 .75.75v6.5h6.5a.75.75 0 0 1 0 1.5h-6.5v6.5a.75.75 0 0 1-1.5 0v-6.5h-6.5a.75.75 0 0 1 0-1.5h6.5v-6.5A.75.75 0 0 1 10 2z"/>
                </svg>
                New email
            </button>
            <div class="ribbon-divider"></div>
            <button class="ribbon-action-btn">Delete</button>
            <button class="ribbon-action-btn">Archive</button>
            <button class="ribbon-action-btn">Sweep</button>
            <button class="ribbon-action-btn">Move to</button>
            <button class="ribbon-action-btn">Categories</button>
        </div>
        <div style="display: flex; gap: 8px; align-items: center;">
            <label style="font-size: 0.8rem; font-weight: 600; color: var(--text-secondary); display: flex; align-items: center; gap: 6px; cursor: pointer; background-color: var(--bg-light); padding: 5px 10px; border-radius: 4px; border: 1px solid var(--gray-border);">
                <input type="checkbox" id="global-auto-reply-toggle" onchange="toggleGlobalAutoReply()" style="cursor: pointer; width: 14px; height: 14px;">
                Auto-Reply Active
            </label>
            <!-- Custom Integrated Toggle Button for our Side Panel -->
            <button class="ribbon-action-btn provident-toggle-btn" onclick="toggleSidePanel()" title="Toggle Provident Operations Copilot Side Panel">
                Operations Copilot
            </button>
        </div>
    </div>

    <!-- 3. MAIN APP BODY -->
    <div class="outlook-body">
        
        <!-- Navigation Rail (Leftmost) -->
        <nav class="outlook-left-rail">
            <div class="left-rail-icon active">
                <svg viewBox="0 0 20 20" width="18" height="18" fill="currentColor">
                    <path d="M17.5 4H2.5A1.5 1.5 0 0 0 1 5.5v9A1.5 1.5 0 0 0 2.5 16h15a1.5 1.5 0 0 0 1.5-1.5v-9A1.5 1.5 0 0 0 17.5 4zm-15 1h15a.5.5 0 0 1 .5.5v1.27l-8 4.8-8-4.8V5.5a.5.5 0 0 1 .5-.5zm15.5 9.5a.5.5 0 0 1-.5.5H2.5a.5.5 0 0 1-.5-.5V7.95l8 4.8 8-4.8z"/>
                </svg>
            </div>
            <div class="left-rail-icon">
                <svg viewBox="0 0 20 20" width="18" height="18" fill="currentColor">
                    <path d="M5.5 2a.5.5 0 0 1 .5.5V3h8v-.5a.5.5 0 0 1 1 0V3h1.5A1.5 1.5 0 0 1 18 4.5v11a1.5 1.5 0 0 1-1.5 1.5H3.5A1.5 1.5 0 0 1 2 15.5v-11A1.5 1.5 0 0 1 3.5 3H5v-.5A.5.5 0 0 1 5.5 2zM3 7v8.5a.5.5 0 0 0 .5.5h13a.5.5 0 0 0 .5-.5V7H3zm13.5-3H3.5a.5.5 0 0 0-.5.5V6h14v-1.5a.5.5 0 0 0-.5-.5z"/>
                </svg>
            </div>
            <div class="left-rail-icon">
                <svg viewBox="0 0 20 20" width="18" height="18" fill="currentColor">
                    <path d="M10 2a3 3 0 1 0 3 3 3 3 0 0 0-3-3zm-4 3a4 4 0 1 1 4 4 4 4 0 0 1-4-4zm-2.5 9A3.5 3.5 0 0 1 7 10.5h6a3.5 3.5 0 0 1 3.5 3.5v1.5a.5.5 0 0 1-.5.5H4a.5.5 0 0 1-.5-.5V14zm1-3.5a2.5 2.5 0 0 0-2.5 2.5v1H15v-1a2.5 2.5 0 0 0-2.5-2.5h-6z"/>
                </svg>
            </div>
            <div class="left-rail-icon">
                <svg viewBox="0 0 20 20" width="18" height="18" fill="currentColor">
                    <path d="M2 4a2 2 0 0 1 2-2h5l2 2h7a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V4zm2 .5V14h12V6.5h-6.5l-2-2H4z"/>
                </svg>
            </div>
            <div class="left-rail-icon">
                <svg viewBox="0 0 20 20" width="18" height="18" fill="currentColor">
                    <path d="M8.2 13.56l-3.3-3.3a.5.5 0 0 1 .7-.7l2.95 2.95 6.95-6.95c.2-.2.5-.2.7 0a.5.5 0 0 1 0 .7l-7.3 7.3a.5.5 0 0 1-.7 0z"/>
                </svg>
            </div>
        </nav>

        <!-- Inbox List (Second Column) -->
        <div class="inbox-list-column">
            <div class="inbox-tabs">
                <div class="inbox-tab active">Focused</div>
                <div class="inbox-tab">Other</div>
            </div>

            <!-- Time Frame Filter -->
            <div class="time-filter-bar" style="padding: 10px 16px; border-bottom: 1px solid var(--gray-border); background-color: #faf9f8; display: flex; flex-direction: column; gap: 8px;">
                <div style="font-size: 0.75rem; font-weight: 600; color: var(--text-secondary); display: flex; justify-content: space-between; align-items: center;">
                    <span>Filter by Timeframe</span>
                    <span id="clear-time-btn" style="color: var(--outlook-blue); cursor: pointer; display: none; font-size: 0.7rem; font-weight: normal;" onclick="clearTimeFilter()">Reset</span>
                </div>
                <div style="display: flex; gap: 6px; align-items: center;">
                    <div style="display: flex; flex-direction: column; flex: 1; gap: 2px;">
                        <span style="font-size: 0.65rem; color: var(--text-muted);">Start</span>
                        <input type="datetime-local" id="filter-start-time" onchange="applyFilters()" style="padding: 3px 6px; border: 1px solid var(--gray-border-dark); border-radius: 3px; font-size: 0.75rem; outline: none; width: 100%;">
                    </div>
                    <div style="display: flex; flex-direction: column; flex: 1; gap: 2px;">
                        <span style="font-size: 0.65rem; color: var(--text-muted);">End</span>
                        <input type="datetime-local" id="filter-end-time" onchange="applyFilters()" style="padding: 3px 6px; border: 1px solid var(--gray-border-dark); border-radius: 3px; font-size: 0.75rem; outline: none; width: 100%;">
                    </div>
                </div>
            </div>
            
            <div class="inbox-scroll">
                <div class="inbox-section-header">
                    <span id="inbox-title">Inbox logs</span>
                    <span class="clear-filter-btn" id="clear-filter-btn" style="display: none;" onclick="clearCategoryFilter()">Clear filter</span>
                </div>
                
                {% for email in emails %}
                <div class="email-card email-item-card" 
                     id="card-{{ email.message_id }}" 
                     data-category="{{ email.category }}"
                     data-timestamp="{{ email.processed_at }}"
                     onclick="loadEmail('{{ email.message_id }}')">
                    <div class="email-avatar" style="background-color: {% if email.category == 'purchase_order' %}#107c41{% elif email.category == 'enquiry' %}#0078d4{% elif email.category == 'invoice' %}#a80000{% elif email.category == 'shipping' %}#8764b8{% else %}#797775{% endif %}">
                        {{ email.sender[:2]|upper }}
                    </div>
                    <div class="email-card-content">
                        <div class="email-card-top">
                            <span class="email-card-sender">{{ email.sender.split('@')[0] }}</span>
                            <span class="email-card-time">14:32</span>
                        </div>
                        <div class="email-card-subject">{{ email.subject or '(No Subject)' }}</div>
                        <div class="email-card-body-snippet">{{ email.body_preview }}</div>
                        <div class="email-card-meta">
                            <span class="outlook-badge {{ email.category }}" id="badge-card-{{ email.message_id }}">
                                {{ email.category|replace('_', ' ') }}
                            </span>
                            {% if loop.index % 2 == 0 %}
                            <span class="attachment-icon-svg">
                                <svg viewBox="0 0 20 20" width="14" height="14" fill="currentColor">
                                    <path d="M14.5 3a3.5 3.5 0 0 0-5 0l-7.1 7.1a2.5 2.5 0 0 0 3.5 3.5l6.75-6.75a1.5 1.5 0 0 0-2.1-2.1L3.9 11.4a.5.5 0 0 0 .7.7l6.6-6.6a.5.5 0 0 1 .7.7L5.3 12.8a1.5 1.5 0 0 1-2.1-2.1l7.1-7.1a2.5 2.5 0 0 1 3.5 3.5l-6.75 6.75a3.5 3.5 0 0 1-5-5L9.3 2.1a4.5 4.5 0 0 1 6.4 6.4l-7.1 7.1a.5.5 0 0 0 .7.7l7.1-7.1a3.5 3.5 0 0 0 0-5z"/>
                                </svg>
                            </span>
                            {% endif %}
                        </div>
                    </div>
                </div>
                {% endfor %}
            </div>
        </div>

        <!-- Central Reading Pane (Third Column) -->
        <div class="outlook-reading-pane">
            <!-- Initial Empty State -->
            <div class="empty-reading-pane" id="reading-pane-empty">
                <svg class="empty-pane-icon-svg" viewBox="0 0 24 24" width="96" height="96" fill="currentColor">
                    <path d="M20 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5l-8-5V6l8 5l8-5v2z"/>
                </svg>
                <div class="empty-pane-title">Select an item to read</div>
                <div class="empty-pane-subtitle">Nothing is selected</div>
            </div>

            <!-- Email Viewer Content -->
            <div class="email-details-view" id="reading-pane-content" style="display: none;">
                <div class="email-view-header">
                    <h2 class="email-view-subject" id="view-subject">Subject line goes here</h2>
                    <div class="email-view-meta-row">
                        <div class="email-view-sender-avatar" id="view-avatar">AZ</div>
                        <div class="email-view-sender-info">
                            <div class="email-view-sender-name" id="view-sender-name">Azure DevOps</div>
                            <div class="email-view-sender-address" id="view-sender-address">noreply@dev.azure.com</div>
                        </div>
                        <div class="email-view-date" id="view-date">Mon 6/29/2026 5:19 PM</div>
                    </div>
                    <div class="email-view-category-row">
                        <span class="outlook-badge" id="view-category-badge">Category</span>
                    </div>
                </div>
                <div class="email-view-body" id="view-body">
                    Email content body text goes here.
                </div>
            </div>
        </div>

        <!-- RESIZABLE DRAG HANDLE -->
        <div class="resize-handle" id="resize-handle"></div>

        <!-- 4. SIDE PANEL (PROVIDENT OPERATIONS COPILOT) -->
        <div class="outlook-side-panel" id="provident-side-panel">
            <div class="panel-header">
                <div class="panel-header-title">
                    <span class="panel-logo">PROVIDENT AI</span>
                </div>
                <button class="close-panel-btn" onclick="toggleSidePanel()">✕</button>
            </div>
            
            <div class="panel-scroll">
                
                <!-- VIEW A: MORNING TRIAGE OVERVIEW (default) -->
                <div id="morning-triage-container" class="triage-widget">
                    <h3 class="side-widget-title">Morning Action Items</h3>
                    <p class="triage-description">Here is your triage checklist for this morning. Click any category to filter your email inbox.</p>
                    
                    <div class="triage-quick-stats" style="margin-bottom: 8px;">
                        <span>Total Processed Today</span>
                        <span id="triage-total-badge">{{ stats.total_processed }}</span>
                    </div>

                    <!-- POs Triage Card -->
                    <div class="triage-checklist-item" data-category="purchase_order" onclick="filterCategory('purchase_order')">
                        <div class="triage-item-bullet" style="background-color: var(--cat-po);"></div>
                        <div class="triage-item-info">
                            <div class="triage-item-title">Purchase Orders</div>
                            <div class="triage-item-sub">Ready to enter in ERP</div>
                        </div>
                        <div class="triage-item-count" id="count-purchase_order">{{ breakdown.purchase_order }}</div>
                    </div>

                    <!-- Enquiry Triage Card -->
                    <div class="triage-checklist-item" data-category="enquiry" onclick="filterCategory('enquiry')">
                        <div class="triage-item-bullet" style="background-color: var(--cat-enquiry);"></div>
                        <div class="triage-item-info">
                            <div class="triage-item-title">Quote Enquiries</div>
                            <div class="triage-item-sub">RFQ details & samples</div>
                        </div>
                        <div class="triage-item-count" id="count-enquiry">{{ breakdown.enquiry }}</div>
                    </div>

                    <!-- Invoice Triage Card -->
                    <div class="triage-checklist-item" data-category="invoice" onclick="filterCategory('invoice')">
                        <div class="triage-item-bullet" style="background-color: var(--cat-invoice);"></div>
                        <div class="triage-item-info">
                            <div class="triage-item-title">Vendor Invoices</div>
                            <div class="triage-item-sub">Verify and route to AP</div>
                        </div>
                        <div class="triage-item-count" id="count-invoice">{{ breakdown.invoice }}</div>
                    </div>

                    <!-- Shipping Triage Card -->
                    <div class="triage-checklist-item" data-category="shipping" onclick="filterCategory('shipping')">
                        <div class="triage-item-bullet" style="background-color: var(--cat-shipping);"></div>
                        <div class="triage-item-info">
                            <div class="triage-item-title">Shipping & Logistics</div>
                            <div class="triage-item-sub">Track dispatch & ETA</div>
                        </div>
                        <div class="triage-item-count" id="count-shipping">{{ breakdown.shipping }}</div>
                    </div>

                    <!-- General Card -->
                    <div class="triage-checklist-item" data-category="general" onclick="filterCategory('general')">
                        <div class="triage-item-bullet" style="background-color: var(--cat-general);"></div>
                        <div class="triage-item-info">
                            <div class="triage-item-title">General Operations</div>
                            <div class="triage-item-sub">Acknowledge & correspondence</div>
                        </div>
                        <div class="triage-item-count" id="count-general">{{ breakdown.general }}</div>
                    </div>
                </div>

                <!-- VIEW B: ACTIVE EMAIL ASSISTANT (visible when email selected) -->
                <div id="active-copilot-container" class="copilot-assistant-view" style="display: none;">
                    <button class="back-triage-btn" onclick="backToTriage()">Back to Triage List</button>
                    
                    <!-- Reclassify Widget -->
                    <div class="copilot-card">
                        <h3 class="copilot-section-title">Email Triage Tag</h3>
                        <select class="reclassify-select" id="reclassify-select" onchange="changeCategory()">
                            <option value="purchase_order">Purchase Order</option>
                            <option value="enquiry">Quote Request / RFQ</option>
                            <option value="invoice">Invoice / Billing</option>
                            <option value="shipping">Shipping / Tracking</option>
                            <option value="general">General Correspondence</option>
                        </select>
                    </div>

                    <!-- Tab headers for Active Assistant widgets -->
                    <div class="copilot-tabs">
                        <div class="copilot-tab active" id="tab-triage" onclick="switchCopilotTab('triage')">Email Summary</div>
                        <div class="copilot-tab" id="tab-catalog" onclick="switchCopilotTab('catalog')">Catalog & Cross-Sell</div>
                        <div class="copilot-tab" id="tab-reply" onclick="switchCopilotTab('reply')">Smart Reply</div>
                    </div>

                    <!-- TAB 1: EMAIL SUMMARY -->
                    <div id="copilot-tab-triage-content">
                        <!-- AI Email Summary Widget -->
                        <div class="copilot-card">
                            <h3 class="copilot-section-title">AI Email Summary</h3>
                            <div id="email-summary-text" style="font-size: 0.84rem; line-height: 1.45; color: var(--text-primary); font-family: Segoe UI, sans-serif; white-space: pre-wrap;">
                                <!-- Populated via Javascript -->
                            </div>
                        </div>

                        <!-- Extracted details widget -->
                        <div class="copilot-card">
                            <h3 class="copilot-section-title">Extracted Details</h3>
                            <div class="extracted-fields-list" id="extracted-fields-container">
                                <!-- Populated via Javascript -->
                            </div>
                        </div>

                        <!-- Pricing Estimator (only visible for enquiries/RFQs) -->
                        <div class="copilot-card" id="pricing-calculator-card" style="display: none;">
                            <h3 class="copilot-section-title">Quotation Pricing Calculator</h3>
                            <div class="triage-description">Adjust parameters below to dynamically calculate box unit cost.</div>
                            <div class="calc-grid">
                                <div class="calc-input-group">
                                    <label for="calc-qty">Quantity</label>
                                    <input type="number" id="calc-qty" value="5000" oninput="calculatePackagingQuote()">
                                </div>
                                <div class="calc-input-group">
                                    <label for="calc-type">Carton Type</label>
                                    <select id="calc-type" onchange="calculatePackagingQuote()">
                                        <option value="single">Single Wall ($0.45)</option>
                                        <option value="double">Double Wall ($0.75)</option>
                                        <option value="heavy" selected>Heavy Duty ($1.20)</option>
                                    </select>
                                </div>
                                <div class="calc-input-group" style="grid-column: span 2; flex-direction: row; align-items: center; justify-content: space-between; padding-top: 6px;">
                                    <label for="calc-print" style="cursor: pointer;">Custom Print Logo Surcharge (+$0.15/unit)</label>
                                    <input type="checkbox" id="calc-print" checked onchange="calculatePackagingQuote()" style="width: 16px; height: 16px;">
                                </div>
                                
                                <div class="calc-result-box">
                                    <div class="calc-total-label">Unit Price:<br><span style="font-size: 0.65rem; color: var(--text-muted);">After Volume Disc.</span></div>
                                    <div class="calc-total-val" id="calc-unit-display">$1.08</div>
                                </div>
                                <div class="calc-result-box" style="margin-top: 0;">
                                    <div class="calc-total-label">Total Quotation Value:</div>
                                    <div class="calc-total-val" id="calc-total-display" style="color: var(--outlook-blue);">$5,400.00</div>
                                </div>

                                <button class="calc-apply-btn" onclick="applyQuoteToDraft()">
                                    Insert Quotation in Draft Reply
                                </button>
                            </div>
                        </div>
                    </div>

                    <!-- TAB 2: CATALOG & CROSS-SELL -->
                    <div id="copilot-tab-catalog-content" style="display: none;">
                        <div class="copilot-card">
                            <h3 class="copilot-section-title">Matched Product Catalog</h3>
                            <div class="catalog-match-box">
                                <span class="match-badge" id="catalog-match-badge">Corrugated</span>
                                <div style="font-size: 0.84rem; font-weight: 600; color: var(--provident-dark);" id="catalog-match-title">Corrugated Boxes / Sheets</div>
                                <div style="font-size: 0.76rem; color: var(--text-secondary); margin-top: 4px;" id="catalog-match-desc">Custom-made shipping and packing boxes. Primary supplier: Fitzpatrick Container Company.</div>
                            </div>
                        </div>

                        <div class="copilot-card">
                            <h3 class="copilot-section-title">Recommended Cross-Sells</h3>
                            <div class="triage-description">Add high-margin complementary packaging products to the quote draft.</div>
                            
                            <div id="cross-sell-list-container">
                                <!-- Populated dynamically -->
                            </div>
                        </div>
                    </div>

                    <!-- TAB 3: SMART REPLY -->
                    <div id="copilot-tab-reply-content" style="display: none;">
                        <div class="copilot-card">
                            <h3 class="copilot-section-title">Suggested Reply Draft</h3>
                            <textarea class="draft-textarea" id="draft-textarea-input" spellcheck="false"></textarea>
                            <div style="display: flex; gap: 8px; margin-top: 6px;">
                                <button class="copy-draft-btn" id="copy-draft-btn" onclick="copyDraftText()" style="flex: 1; margin-top: 0; display: flex; align-items: center; justify-content: center; gap: 6px;">
                                    Copy Draft
                                </button>
                                <button class="copy-draft-btn" id="send-reply-btn" onclick="sendReply()" style="flex: 1; margin-top: 0; background-color: var(--cat-po); display: flex; align-items: center; justify-content: center; gap: 6px;">
                                    Send Reply
                                </button>
                            </div>
                        </div>
                    </div>
                </div>

            </div>
        </div>

    </div>

    <!-- 5. SCRIPTS -->
    <script>
        // Store emails object in javascript for easy reading pane rendering
        const emailsMap = {};
        {% for email in emails %}
        emailsMap["{{ email.message_id }}"] = {
            messageId: {{ email.message_id|tojson }},
            sender: {{ email.sender|tojson }},
            subject: {{ email.subject|tojson }},
            body: {{ email.body_preview|tojson }},
            category: {{ email.category|tojson }},
            confidence: {{ email.confidence|tojson }},
            reason: {{ email.reason|tojson }},
            processedAt: {{ email.processed_at|tojson }},
            extractedData: {{ email.extracted_data|safe if email.extracted_data else '{}' }},
            responseDraft: {{ email.response_draft|tojson if email.response_draft else '""' }},
            email_status: {{ email.email_status|tojson if email.email_status else '""' }},
            reply_sent: {{ email.reply_sent|tojson if email.reply_sent else 'false' }},
            sent_reply: {{ email.sent_reply|tojson if email.sent_reply else '""' }}
        };
        {% endfor %}

        // Resizable Sidebar JavaScript Implementation
        const handle = document.getElementById('resize-handle');
        const panel = document.getElementById('provident-side-panel');
        let isResizing = false;

        handle.addEventListener('mousedown', (e) => {
            isResizing = true;
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
            handle.classList.add('active');
        });

        document.addEventListener('mousemove', (e) => {
            if (!isResizing) return;
            const newWidth = window.innerWidth - e.clientX;
            // Bound resizing width between 280px and 600px
            if (newWidth >= 280 && newWidth <= 600) {
                panel.style.width = newWidth + 'px';
            }
        });

        document.addEventListener('mouseup', () => {
            if (isResizing) {
                isResizing = false;
                document.body.style.cursor = 'default';
                document.body.style.userSelect = 'auto';
                handle.classList.remove('active');
            }
        });

        // Sidebar View Toggle
        let selectedMessageId = null;

        function toggleSidePanel() {
            panel.classList.toggle('collapsed');
            if (panel.classList.contains('collapsed')) {
                handle.style.display = 'none';
            } else {
                handle.style.display = 'block';
            }
        }

        // Triage Category Filter JavaScript
        let currentFilterCategory = null;

        function applyFilters() {
            const startVal = document.getElementById('filter-start-time').value;
            const endVal = document.getElementById('filter-end-time').value;
            
            const hasStart = startVal !== "";
            const hasEnd = endVal !== "";
            
            const startMs = hasStart ? new Date(startVal).getTime() : 0;
            const endMs = hasEnd ? new Date(endVal).getTime() : Infinity;
            
            if (hasStart || hasEnd) {
                document.getElementById('clear-time-btn').style.display = 'inline';
            } else {
                document.getElementById('clear-time-btn').style.display = 'none';
            }

            document.querySelectorAll('.email-item-card').forEach(card => {
                const matchesCategory = !currentFilterCategory || (card.dataset.category === currentFilterCategory);
                
                const timestampStr = card.dataset.timestamp;
                const timeMs = new Date(timestampStr).getTime();
                const matchesTime = (timeMs >= startMs && timeMs <= endMs);
                
                if (matchesCategory && matchesTime) {
                    card.style.display = 'flex';
                } else {
                    card.style.display = 'none';
                }
            });
        }

        function clearTimeFilter() {
            document.getElementById('filter-start-time').value = '';
            document.getElementById('filter-end-time').value = '';
            document.getElementById('clear-time-btn').style.display = 'none';
            applyFilters();
        }

        function filterCategory(category) {
            if (currentFilterCategory === category) {
                clearCategoryFilter();
                return;
            }

            currentFilterCategory = category;
            
            // Highlight selected filter item in sidebar
            document.querySelectorAll('.triage-checklist-item').forEach(item => {
                if (item.dataset.category === category) {
                    item.classList.add('active-filter');
                } else {
                    item.classList.remove('active-filter');
                }
            });

            // Show Clear Filter indicator
            document.getElementById('clear-filter-btn').style.display = 'inline';
            document.getElementById('inbox-title').innerText = category.replace('_', ' ') + 's';
            
            applyFilters();
        }

        function clearCategoryFilter() {
            currentFilterCategory = null;
            
            // Remove highlighted filter item
            document.querySelectorAll('.triage-checklist-item').forEach(item => {
                item.classList.remove('active-filter');
            });

            // Hide Clear Filter indicator
            document.getElementById('clear-filter-btn').style.display = 'none';
            document.getElementById('inbox-title').innerText = 'Inbox logs';
            
            applyFilters();
        }

        // Email Reader Loader
        function loadEmail(msgId) {
            selectedMessageId = msgId;

            // Remove selection from all cards
            document.querySelectorAll('.email-item-card').forEach(card => {
                card.classList.remove('selected');
            });
            
            // Mark clicked card as read & selected
            const selectedCard = document.getElementById('card-' + msgId);
            if (selectedCard) {
                selectedCard.classList.add('selected');
                selectedCard.classList.remove('unread');
            }

            const email = emailsMap[msgId];
            if (!email) return;

            // Render email details in central reading pane
            document.getElementById('reading-pane-empty').style.display = 'none';
            const contentPane = document.getElementById('reading-pane-content');
            contentPane.style.display = 'flex';

            document.getElementById('view-subject').innerText = email.subject || '(No Subject)';
            document.getElementById('view-sender-name').innerText = email.sender.split('@')[0];
            document.getElementById('view-sender-address').innerText = email.sender;
            document.getElementById('view-avatar').innerText = email.sender.slice(0, 2).toUpperCase();
            
            // Render category badge
            const badge = document.getElementById('view-category-badge');
            badge.className = 'outlook-badge ' + email.category;
            badge.innerText = email.category.replace('_', ' ');
            
            // Set sender avatar background color based on category
            const avatar = document.getElementById('view-avatar');
            let color = '#797775';
            if (email.category === 'purchase_order') color = '#107c41';
            else if (email.category === 'enquiry') color = '#0078d4';
            else if (email.category === 'invoice') color = '#a80000';
            else if (email.category === 'shipping') color = '#8764b8';
            avatar.style.backgroundColor = color;

            // Escape helper
            function esc(str) {
                return String(str || '')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
            }

            // Render full detailed email body text
            let fullBodyText = email.body;
            if (email.category === 'purchase_order') {
                fullBodyText = `Hi Team,\n\nPlease process our order (PO #${email.extractedData['PO Number'] || '99231'}) for ${email.extractedData['Quantity'] || '2,500 units'} of ${email.extractedData['Items Ordered'] || 'boxes'}.\n\nWe require delivery to our main warehouse by ${email.extractedData['Delivery Date'] || 'July 12, 2026'}.\n\nLet us know if you need anything else.\n\nBest regards,\n${email.sender.split('@')[0]}\nProcurement Manager`;
            } else if (email.category === 'enquiry') {
                fullBodyText = `Dear Sales,\n\nWe are looking to source ${email.extractedData['Quantity'] || '5,000'} pieces of ${email.extractedData['Requested Item'] || 'Heavy Duty Cartons'}. Could you please send us a cost estimate?\n\n- Custom Printing: ${email.extractedData['Custom Printing (Yes/No)'] || 'Yes'}\n- Samples required: ${email.extractedData['Samples Requested (Yes/No)'] || 'Yes'}\n\nThis is an ${email.extractedData['Urgency'] || 'Medium'} urgency request. Thank you,\n${email.sender.split('@')[0]}`;
            } else if (email.category === 'invoice') {
                fullBodyText = `Hi Billing,\n\nAttached is invoice ${email.extractedData['Invoice Number'] || 'INV-2026-4401'} from ${email.extractedData['Vendor'] || 'Logistics'}.\n\n- Amount Due: ${email.extractedData['Amount Due'] || '$3,450.00'}\n- Payment Due Date: ${email.extractedData['Due Date'] || 'July 29, 2026'}\n- Payment Terms: ${email.extractedData['Payment Terms'] || 'Net 30'}\n\nPlease verify and settle. Thanks,\nAccounting Team`;
            } else if (email.category === 'shipping') {
                fullBodyText = `Hello Ops,\n\nYour freight shipment has been dispatched. Details below:\n\n- Carrier: ${email.extractedData['Carrier'] || 'DHL'}\n- Waybill / Tracking: ${email.extractedData['Tracking Number'] || '88201882'}\n- Estimated Arrival: ${email.extractedData['Est Delivery Date'] || 'Wednesday, July 1'}\n- Current Status: ${email.extractedData['Current Status'] || 'In Transit'}\n\nLogistics Dispatch Service`;
            }

            let bodyHtml = `<div style="white-space: pre-wrap; font-family: Segoe UI, Arial, sans-serif;">${esc(fullBodyText)}</div>`;
            if (email.reply_sent || email.sent_reply) {
                const sentText = email.sent_reply || email.responseDraft || "Thank you for contacting us. We have received your email and are processing it.";
                bodyHtml += `
                    <div style="margin-top: 24px; border-top: 1px solid var(--gray-border); padding-top: 16px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 0.8rem; color: var(--text-secondary);">
                            <strong>Your Reply (Sent via Auto-Reply/Copilot)</strong>
                            <span>Sent</span>
                        </div>
                        <div style="padding: 12px; background-color: var(--bg-light); border-left: 3px solid var(--outlook-blue); border-radius: 4px; white-space: pre-wrap; font-family: Arial, Helvetica, sans-serif; font-size: 0.9rem; color: var(--text-primary); line-height: 1.45;">${esc(sentText)}</div>
                    </div>
                `;
            }
            document.getElementById('view-body').innerHTML = bodyHtml;

            // Render AI Email Summary Text
            document.getElementById('email-summary-text').innerText = email.reason || 'No summary available.';

            // Render AI Copilot Sidebar view (View B)
            document.getElementById('morning-triage-container').style.display = 'none';
            document.getElementById('active-copilot-container').style.display = 'flex';

            // Select active option in Reclassify dropdown
            document.getElementById('reclassify-select').value = email.category;

            // Reset Copilot tabs to first tab "Email Summary"
            switchCopilotTab('triage');

            // Render Extracted Metadata Fields
            const fieldsContainer = document.getElementById('extracted-fields-container');
            fieldsContainer.innerHTML = '';
            
            const dataObj = email.extractedData;
            const keys = Object.keys(dataObj);
            
            if (keys.length > 0) {
                keys.forEach(k => {
                    const row = document.createElement('div');
                    row.className = 'extracted-field-row';
                    row.innerHTML = `
                        <div class="field-label">${k}</div>
                        <div class="field-value-container">
                            <span class="field-value">${dataObj[k]}</span>
                            <button class="copy-btn" onclick="copyToClipboard('${dataObj[k].replace(/'/g, "\\'")}', this)" title="Copy field">
                                <svg viewBox="0 0 16 16" width="12" height="12" fill="currentColor" style="margin-right: 4px;">
                                    <path d="M4 1.5H3a2 2 0 0 0-2 2V12a2 2 0 0 0 2 2h7a2 2 0 0 0 2-2v-1h-1v1a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3.5a1 1 0 0 1 1-1h1v-1z"/>
                                    <path d="M9.5 1a1.5 1.5 0 0 1 1.5 1.5v9A1.5 1.5 0 0 1 9.5 13h-5A1.5 1.5 0 0 1 3 11.5v-9A1.5 1.5 0 0 1 4.5 1h5zm0 1h-5a.5.5 0 0 0-.5.5v9a.5.5 0 0 0 .5.5h5a.5.5 0 0 0 .5-.5v-9a.5.5 0 0 0-.5-.5z"/>
                                </svg>
                                Copy
                            </button>
                        </div>
                    `;
                    fieldsContainer.appendChild(row);
                });
            } else {
                fieldsContainer.innerHTML = '<div style="font-size:0.8rem; color:var(--text-muted); text-align:center; padding: 10px 0;">No details extracted.</div>';
            }

            // Render Tab-specific conditional blocks (Pricing Calculator)
            const pricingCard = document.getElementById('pricing-calculator-card');
            pricingCard.style.display = 'none';

            if (email.category === 'enquiry') {
                pricingCard.style.display = 'block';
                // Pre-populate pricing inputs
                const parsedQty = parseInt(email.extractedData['Quantity']) || 5000;
                document.getElementById('calc-qty').value = parsedQty;
                
                // Select type
                const itemRequested = (email.extractedData['Requested Item'] || '').toLowerCase();
                const typeSelect = document.getElementById('calc-type');
                if (itemRequested.includes('heavy') || itemRequested.includes('5-ply')) {
                    typeSelect.value = 'heavy';
                } else if (itemRequested.includes('double') || itemRequested.includes('3-ply')) {
                    typeSelect.value = 'double';
                } else {
                    typeSelect.value = 'single';
                }

                // Check custom print
                const printReq = (email.extractedData['Custom Printing (Yes/No)'] || '').toLowerCase();
                document.getElementById('calc-print').checked = (printReq.includes('yes') || printReq.includes('logo'));
                
                calculatePackagingQuote();
            }

            // Populate Catalog & Cross-Sell Tab details
            renderCatalogAndCrossSells(email);

            // Render Suggested Draft reply text
            document.getElementById('draft-textarea-input').value = email.responseDraft;
            
            // Make sure the side panel is expanded
            panel.classList.remove('collapsed');
            handle.style.display = 'block';
        }

        function backToTriage() {
            selectedMessageId = null;
            document.getElementById('active-copilot-container').style.display = 'none';
            document.getElementById('morning-triage-container').style.display = 'flex';
        }

        // Active Copilot Tab Switcher
        function switchCopilotTab(tabName) {
            // Remove active classes
            document.querySelectorAll('.copilot-tab').forEach(t => t.classList.remove('active'));
            document.getElementById('copilot-tab-triage-content').style.display = 'none';
            document.getElementById('copilot-tab-catalog-content').style.display = 'none';
            document.getElementById('copilot-tab-reply-content').style.display = 'none';

            // Set active class
            document.getElementById('tab-' + tabName).classList.add('active');
            document.getElementById('copilot-tab-' + tabName + '-content').style.display = 'block';
        }

        // Render matched catalog details and cross sells
        function renderCatalogAndCrossSells(email) {
            const matchBadge = document.getElementById('catalog-match-badge');
            const matchTitle = document.getElementById('catalog-match-title');
            const matchDesc = document.getElementById('catalog-match-desc');
            const crossContainer = document.getElementById('cross-sell-list-container');
            
            crossContainer.innerHTML = '';
            
            let crossSells = [];
            
            if (email.category === 'purchase_order' || email.category === 'enquiry') {
                matchBadge.innerText = 'Corrugated';
                matchBadge.style.backgroundColor = 'var(--cat-po)';
                matchTitle.innerText = 'Corrugated Boxes & Folders';
                matchDesc.innerText = 'Philly-manufactured corrugated boxes, sheets, and folders. Primary Supplier: Fitzpatrick Container Company.';
                
                crossSells = [
                    { title: "IPG Water-Activated Tape (Gummed)", vendor: "Intertape Polymer Group", price: "$18.50/roll", id: "tape" },
                    { title: "Pregis Air-Pillow Bubble Wrap (1/2-inch)", vendor: "Pregis Corporation", price: "$42.00/roll", id: "bubble" },
                    { title: "AEP Stretch Wrap Film (80 gauge)", vendor: "AEP Industries", price: "$34.00/roll", id: "film" }
                ];
            } else if (email.category === 'invoice') {
                matchBadge.innerText = 'Shipping';
                matchBadge.style.backgroundColor = 'var(--provident-gold)';
                matchTitle.innerText = 'Freight & Logistics Services';
                matchDesc.innerText = 'Strategic inbound/outbound freight management with mid-atlantic shipping partners.';
                
                crossSells = [
                    { title: "Corner Boards / Edge Protectors", vendor: "PAC Strapping", price: "$2.10/each", id: "boards" },
                    { title: "Polypropylene Strapping Reels", vendor: "PAC Strapping", price: "$98.00/reel", id: "straps" }
                ];
            } else {
                matchBadge.innerText = 'Janitorial';
                matchBadge.style.backgroundColor = 'var(--text-secondary)';
                matchTitle.innerText = 'Facility & Industrial Supplies';
                matchDesc.innerText = 'FDA-approved food-service packaging, building maintenance, and paper towels.';
                
                crossSells = [
                    { title: "Heavy Duty Trash Liners (55 Gal)", vendor: "Bay West Paper", price: "$48.00/case", id: "liners" },
                    { title: "FDA Latex Gloves & Hairnets", vendor: "Food Safety Supplies", price: "$14.50/box", id: "gloves" }
                ];
            }
            
            crossSells.forEach(item => {
                const row = document.createElement('div');
                row.className = 'cross-sell-item';
                row.innerHTML = `
                    <div class="cross-sell-info">
                        <div class="cross-sell-title">${item.title}</div>
                        <div class="cross-sell-vendor">${item.vendor} &bull; ${item.price}</div>
                    </div>
                    <button class="add-cross-btn" id="btn-cross-${item.id}" onclick="addCrossSell('${item.title.replace(/'/g, "\\'")}', '${item.id}')">+</button>
                `;
                crossContainer.appendChild(row);
            });
        }

        // Add Cross Sell description to Draft response
        function addCrossSell(itemName, itemId) {
            const btn = document.getElementById('btn-cross-' + itemId);
            if (btn.classList.contains('added')) return;
            
            btn.classList.add('added');
            btn.innerText = '✔';
            
            // Append product offer to the draft reply text
            const draftInput = document.getElementById('draft-textarea-input');
            let currentText = draftInput.value;
            
            // Insert cross sell before the signoff
            const signoffIndex = currentText.toLowerCase().lastIndexOf('best regards');
            const crossSellMessage = `\\nAlso, we noticed you might need packing supplies for this dispatch. We currently have stock of ${itemName}. Let me know if you would like me to add this to your invoice.\\n`;
            
            if (signoffIndex !== -1) {
                currentText = currentText.substring(0, signoffIndex) + crossSellMessage + "\\n" + currentText.substring(signoffIndex);
            } else {
                currentText += "\\n" + crossSellMessage;
            }
            
            draftInput.value = currentText;
            
            // Update local memory map
            emailsMap[selectedMessageId].responseDraft = currentText;
        }

        // Pricing Calculator Logic
        function calculatePackagingQuote() {
            const qty = parseInt(document.getElementById('calc-qty').value) || 0;
            const cartonType = document.getElementById('calc-type').value;
            const printCustom = document.getElementById('calc-print').checked;

            let baseCost = 0.45;
            if (cartonType === 'double') baseCost = 0.75;
            else if (cartonType === 'heavy') baseCost = 1.20;

            let printSurcharge = printCustom ? 0.15 : 0;
            let unitPrice = baseCost + printSurcharge;

            // Apply Volume Discounts
            let discount = 0;
            if (qty >= 5000) discount = 0.20; // 20% off
            else if (qty >= 1000) discount = 0.10; // 10% off

            unitPrice = unitPrice * (1 - discount);
            const totalVal = unitPrice * qty;

            document.getElementById('calc-unit-display').innerText = '$' + unitPrice.toFixed(2);
            document.getElementById('calc-total-display').innerText = '$' + totalVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }

        function applyQuoteToDraft() {
            if (!selectedMessageId) return;
            
            const unitPrice = document.getElementById('calc-unit-display').innerText;
            const totalVal = document.getElementById('calc-total-display').innerText;
            const qty = document.getElementById('calc-qty').value;
            const cartonLabel = document.getElementById('calc-type').options[document.getElementById('calc-type').selectedIndex].text.split(' ($')[0];

            const email = emailsMap[selectedMessageId];
            
            // Create professional quote reply text
            const quoteDraft = `Hi ${email.sender.split('@')[0]},\n\nThank you for your inquiry. We are pleased to provide the custom manufacturing quotation for your packaging requirements:\n\nQuotation Details:\n- Item: ${cartonLabel}\n- Quantity: ${parseInt(qty).toLocaleString()} units\n- Unit Price: ${unitPrice} per carton\n- Total Estimated Order Value: ${totalVal} (FOB Philadelphia Warehouse)\n- Estimated Production Lead Time: 5-7 business days\n\nIf the quotation matches your specifications, please reply directly and we will generate the invoice and order contract.\n\nBest regards,\nProvident Packaging Sales Team\nPhiladelphia, PA\n(215) 827-0960`;

            document.getElementById('draft-textarea-input').value = quoteDraft;
            emailsMap[selectedMessageId].responseDraft = quoteDraft;
            
            // Switch tab to Smart Reply to show the updated draft!
            switchCopilotTab('reply');
        }

        // Send Reply via API
        function sendReply() {
            if (!selectedMessageId) return;
            const replyText = document.getElementById('draft-textarea-input').value;
            const btn = document.getElementById('send-reply-btn');
            
            btn.innerText = 'Sending Reply...';
            btn.disabled = true;

            fetch('/api/send_reply', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    message_id: selectedMessageId,
                    reply_text: replyText
                })
            })
            .then(response => response.json())
            .then(data => {
                btn.innerText = 'Send Reply';
                btn.disabled = false;
                if (data.success) {
                    // Update category status and reply state in local mapping
                    emailsMap[selectedMessageId].email_status = data.new_status;
                    emailsMap[selectedMessageId].reply_sent = true;
                    emailsMap[selectedMessageId].sent_reply = replyText;
                    
                    // Reload email view (appends reply and advances stepper)
                    loadEmail(selectedMessageId);
                    
                    // Update Morning triage counts
                    updateTriageCounts();
                    
                    alert('Reply registered successfully! Stepper advanced.');
                } else {
                    alert('Failed to send reply: ' + data.error);
                }
            })
            .catch(error => {
                btn.innerText = 'Send Reply';
                btn.disabled = false;
                console.error('Error sending reply:', error);
                alert('Connection error sending reply.');
            });
        }

        // Global Auto-Reply Toggles
        function toggleGlobalAutoReply() {
            const isActive = document.getElementById('global-auto-reply-toggle').checked;
            fetch('/api/settings', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ auto_reply: isActive })
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    console.log('Auto-Reply setting updated:', data.auto_reply);
                }
            })
            .catch(error => console.error('Error updating settings:', error));
        }

        function loadAutoReplySetting() {
            fetch('/api/settings')
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    document.getElementById('global-auto-reply-toggle').checked = data.auto_reply;
                }
            })
            .catch(error => console.error('Error loading settings:', error));
        }

        // Change Category dropdown API callback
        function changeCategory() {
            if (!selectedMessageId) return;
            const newCat = document.getElementById('reclassify-select').value;
            
            fetch('/api/reclassify', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    message_id: selectedMessageId,
                    category: newCat
                })
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    // Update category in local memory
                    emailsMap[selectedMessageId].category = newCat;
                    
                    // Update Outlook list card badge and class
                    const card = document.getElementById('card-' + selectedMessageId);
                    card.dataset.category = newCat;
                    
                    const badgeCard = document.getElementById('badge-card-' + selectedMessageId);
                    badgeCard.className = 'outlook-badge ' + newCat;
                    badgeCard.innerText = newCat.replace('_', ' ');
                    
                    // Update selected avatar color
                    const avatar = card.querySelector('.email-avatar');
                    let color = '#797775';
                    if (newCat === 'purchase_order') color = '#107c41';
                    else if (newCat === 'enquiry') color = '#0078d4';
                    else if (newCat === 'invoice') color = '#a80000';
                    else if (newCat === 'shipping') color = '#8764b8';
                    avatar.style.backgroundColor = color;

                    // Update main email detail badge
                    const viewBadge = document.getElementById('view-category-badge');
                    viewBadge.className = 'outlook-badge ' + newCat;
                    viewBadge.innerText = newCat.replace('_', ' ');
                    document.getElementById('view-avatar').style.backgroundColor = color;

                    // Refresh tabs and catalog match display
                    loadEmail(selectedMessageId);

                    // Update counts in morning triage list
                    updateTriageCounts();
                } else {
                    alert('Failed to update category: ' + data.error);
                }
            })
            .catch(error => {
                console.error('Error reclassifying:', error);
                alert('Connection error changing category.');
            });
        }

        // Update counts dynamically in morning triage checklist
        function updateTriageCounts() {
            const counts = {
                purchase_order: 0,
                enquiry: 0,
                invoice: 0,
                shipping: 0,
                general: 0
            };
            
            Object.values(emailsMap).forEach(e => {
                if (e.category in counts) {
                    counts[e.category]++;
                }
            });
            
            document.getElementById('count-purchase_order').innerText = counts.purchase_order;
            document.getElementById('count-enquiry').innerText = counts.enquiry;
            document.getElementById('count-invoice').innerText = counts.invoice;
            document.getElementById('count-shipping').innerText = counts.shipping;
            document.getElementById('count-general').innerText = counts.general;
        }

        // Copy actions
        function copyToClipboard(text, element) {
            navigator.clipboard.writeText(text).then(() => {
                const originalHTML = element.innerHTML;
                element.innerHTML = '✔ Copied';
                element.classList.add('copied');
                setTimeout(() => {
                    element.innerHTML = originalHTML;
                    element.classList.remove('copied');
                }, 1500);
            });
        }

        function copyDraftText() {
            const text = document.getElementById('draft-textarea-input').value;
            const element = document.getElementById('copy-draft-btn');
            
            navigator.clipboard.writeText(text).then(() => {
                const originalText = element.innerHTML;
                element.innerHTML = '✔ Suggested Response Copied!';
                element.classList.add('copied');
                setTimeout(() => {
                    element.innerHTML = originalText;
                    element.classList.remove('copied');
                }, 2000);
            });
        }

        // Initialize settings on load
        loadAutoReplySetting();
    </script>
</body>
</html>
"""

@app.route("/")
def home():
    # Fetch data
    stats = db.get_accuracy()
    emails = db.get_recent_emails(50)
    
    # Process breakdown
    breakdown = {
        "purchase_order": 0,
        "enquiry": 0,
        "invoice": 0,
        "shipping": 0,
        "general": 0
    }
    
    for e in emails:
        cat = e["category"]
        if cat in breakdown:
            breakdown[cat] += 1
            
    total = len(emails)
    percentage = {cat: round((count / total * 100), 1) if total > 0 else 0 for cat, count in breakdown.items()}
    
    # Process daily stats for chart
    daily_raw = db.get_daily_stats(days=7)
    
    # Group daily raw by date
    daily_grouped = {}
    for d in daily_raw:
        date = d["date"]
        daily_grouped[date] = daily_grouped.get(date, 0) + d["count"]
        
    # Sort chronologically
    sorted_dates = sorted(daily_grouped.keys())
    chart_data = None
    if sorted_dates:
        chart_data = {
            "labels": sorted_dates,
            "values": [daily_grouped[date] for date in sorted_dates]
        }
        
    return render_template_string(
        DASHBOARD_HTML,
        stats=stats,
        emails=emails,
        breakdown=breakdown,
        percentage=percentage,
        chart_data=chart_data
    )

@app.route("/api/feedback", methods=["POST"])
def save_feedback():
    data = request.get_json() or {}
    msg_id = data.get("message_id")
    is_correct = data.get("is_correct")
    
    if not msg_id or is_correct is None:
        return jsonify({"success": False, "error": "Missing parameters"}), 400
        
    result = db.add_feedback(msg_id, is_correct)
    if result:
        # Fetch updated statistics to return to UI
        stats = db.get_accuracy()
        return jsonify({"success": True, "stats": stats})
    else:
        return jsonify({"success": False, "error": "Email record not found"}), 404

@app.route("/api/reclassify", methods=["POST"])
def reclassify():
    data = request.get_json() or {}
    msg_id = data.get("message_id")
    category = data.get("category")
    
    if not msg_id or not category:
        return jsonify({"success": False, "error": "Missing parameters"}), 400
        
    result = db.update_category(msg_id, category)
    if result:
        # Fetch updated statistics to return to UI
        stats = db.get_accuracy()
        return jsonify({"success": True, "stats": stats})
    else:
        return jsonify({"success": False, "error": "Email record not found"}), 404

@app.route("/api/send_reply", methods=["POST"])
def send_reply():
    data = request.get_json() or {}
    msg_id = data.get("message_id")
    reply_text = data.get("reply_text", "")
    
    if not msg_id:
        return jsonify({"success": False, "error": "Missing message_id"}), 400
        
    new_status = db.record_reply_sent(msg_id, reply_text)
    if new_status:
        return jsonify({"success": True, "new_status": new_status})
    else:
        return jsonify({"success": False, "error": "Email record not found"}), 404

@app.route("/api/settings", methods=["GET", "POST"])
def settings_endpoint():
    if request.method == "POST":
        data = request.get_json() or {}
        auto_reply = data.get("auto_reply", False)
        config.AUTO_REPLY_ENABLED = auto_reply
        return jsonify({"success": True, "auto_reply": config.AUTO_REPLY_ENABLED})
    else:
        return jsonify({"success": True, "auto_reply": getattr(config, "AUTO_REPLY_ENABLED", False)})

# Serve Add-in Static Files
@app.route("/addin/<path:filename>")
def serve_addin(filename):
    addin_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outlook-addin")
    return send_from_directory(addin_dir, filename)

# Dedicated manifest route with explicit XML content-type (required for "Install from URL")
@app.route("/manifest.xml")
def serve_manifest():
    addin_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outlook-addin")
    response = send_from_directory(addin_dir, "manifest.xml")
    response.headers["Content-Type"] = "application/xml; charset=utf-8"
    return response

# Add-in API: Classify email on demand (or fetch if already cached)
@app.route("/api/addin/classify", methods=["POST"])
def addin_classify():
    data = request.get_json() or {}
    message_id = data.get("message_id")
    subject = data.get("subject", "")
    body = data.get("body", "")
    sender = data.get("sender", "")
    conversation_id = data.get("conversation_id")
    source_folder = data.get("source_folder", "Inbox")

    if not message_id:
        return jsonify({"success": False, "error": "Missing message_id parameter"}), 400

    try:
        # Check database
        session = db.Session()
        from database import ProcessedEmail
        email_record = session.query(ProcessedEmail).filter_by(message_id=message_id).first()
        
        if email_record:
            response_data = {
                "message_id": email_record.message_id,
                "category": email_record.category,
                "confidence": email_record.confidence,
                "reason": email_record.reason,
                "extracted_data": email_record.extracted_data,
                "response_draft": email_record.response_draft,
                "email_status": email_record.email_status,
                "estimated_value": email_record.estimated_value,
                "priority_score": email_record.priority_score,
                "priority_tier": email_record.priority_tier,
                "source_folder": email_record.source_folder or "Inbox",
                "processed_at": email_record.processed_at.isoformat() if email_record.processed_at else None
            }
            session.close()
            return jsonify({"success": True, "data": response_data})

        session.close()
        
        classifier = EmailClassifier()
        result = classifier.classify(subject, body)
        
        category = result["category"]
        confidence = result["confidence"]
        reason = result["reason"]
        extracted = result["extracted_data"]
        draft = result["response_draft"]
        estimated_value = result.get("estimated_value", 0.0)
        
        outlook_cat = config.CATEGORIES.get(category, {}).get("outlook_category", "General")
        
        from datetime import datetime
        from priority import compute_priority_score
        from database import INITIAL_STATUS
        
        # Compute priority before saving
        priority_result = compute_priority_score({
            "received_at": datetime.utcnow().isoformat(),
            "estimated_value": estimated_value,
            "subject": subject,
            "body_preview": body[:200],
            "sender": sender,
            "category": category,
        })
        
        db.record_classification(
            message_id=message_id,
            user_email="info@providentpackaging.com",
            sender=sender,
            subject=subject,
            body_preview=body[:200],
            category=category,
            confidence=confidence,
            reason=reason,
            outlook_category=outlook_cat,
            received_at=datetime.utcnow(),
            extracted_data=extracted,
            response_draft=draft,
            conversation_id=conversation_id,
            estimated_value=estimated_value,
            source_folder=source_folder,
        )
        db.update_priority(message_id, priority_result["score"], priority_result["tier"])
        
        response_data = {
            "message_id": message_id,
            "category": category,
            "confidence": confidence,
            "reason": reason,
            "extracted_data": extracted,
            "response_draft": draft,
            "email_status": INITIAL_STATUS.get(category, "gen_new"),
            "estimated_value": estimated_value,
            "priority_score": priority_result["score"],
            "priority_tier": priority_result["tier"],
            "priority_reasons": priority_result["reasons"],
            "source_folder": source_folder,
            "processed_at": datetime.utcnow().isoformat()
        }
        
        return jsonify({"success": True, "data": response_data})
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

# Add-in API: Override/reclassify category
@app.route("/api/addin/reclassify", methods=["POST"])
def addin_reclassify():
    data = request.get_json() or {}
    message_id = data.get("message_id")
    category = data.get("category")

    if not message_id or not category:
        return jsonify({"success": False, "error": "Missing parameters"}), 400

    try:
        # Check if record exists
        session = db.Session()
        from database import ProcessedEmail
        email_record = session.query(ProcessedEmail).filter_by(message_id=message_id).first()
        
        if not email_record:
            # Create a shell record if it doesn't exist
            from datetime import datetime
            session.close()
            
            outlook_cat = config.CATEGORIES.get(category, {}).get("outlook_category", "General")
            db.record_classification(
                message_id=message_id,
                user_email="info@providentpackaging.com",
                sender="unknown@providentpackaging.com",
                subject="Manual Override Mail",
                body_preview="Manual Override Mail",
                category=category,
                confidence=100.0,
                reason="User manually set category in Outlook Add-in",
                outlook_category=outlook_cat,
                received_at=datetime.utcnow()
            )
            return jsonify({"success": True})
            
        session.close()
        
        # Update category in DB
        result = db.update_category(message_id, category)
        if result:
            return jsonify({"success": True})
        else:
            return jsonify({"success": False, "error": "Failed to update category"}), 500
            
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/addin/triage_summary", methods=["GET"])
def addin_triage_summary():
    try:
        session = db.Session()
        from database import ProcessedEmail
        from sqlalchemy import func
        from datetime import datetime, date
        from priority import compute_priority_score

        today = date.today()
        results = session.query(
            ProcessedEmail.category, func.count(ProcessedEmail.id)
        ).filter(
            ProcessedEmail.processed_at >= datetime(today.year, today.month, today.day)
        ).group_by(ProcessedEmail.category).all()
        session.close()

        counts = {"purchase_order": 0, "enquiry": 0, "invoice": 0, "shipping": 0, "general": 0}
        total = 0
        for cat, count in results:
            if cat in counts:
                counts[cat] = count
                total += count
            else:
                counts["general"] += count
                total += count

        # Count critical-priority items per category using live priority scoring
        critical_counts = {k: 0 for k in counts}
        all_emails = db.get_recent_emails(limit=200)
        for e in all_emails:
            cat = e.get("category", "general")
            if cat in critical_counts:
                p = compute_priority_score(e)
                if p["tier"] == "critical":
                    critical_counts[cat] = critical_counts.get(cat, 0) + 1

        return jsonify({
            "success": True,
            "counts": counts,
            "critical_counts": critical_counts,
            "total": total
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/addin/category_emails", methods=["GET"])
def addin_category_emails():
    try:
        category = request.args.get("category")
        priority_filter = request.args.get("priority")  # optional: critical/high/medium/low
        status_filter = request.args.get("status")      # optional: specific status ID
        if not category:
            return jsonify({"success": False, "error": "Category is required"}), 400

        from priority import compute_batch_priorities
        emails = db.get_emails_by_category(category, limit=100)
        emails = compute_batch_priorities(emails)

        # Save computed priorities back to DB
        for e in emails:
            if e.get("priority_score") is not None:
                db.update_priority(e["message_id"], e["priority_score"], e["priority_tier"])

        # Apply optional filters
        if priority_filter:
            emails = [e for e in emails if e.get("priority_tier") == priority_filter]
        if status_filter:
            emails = [e for e in emails if e.get("email_status") == status_filter]

        return jsonify({"success": True, "emails": emails})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/addin/update_status", methods=["POST"])
def addin_update_status():
    """Manually advance or set lifecycle status for an email."""
    data = request.get_json() or {}
    message_id = data.get("message_id")
    new_status = data.get("status")
    if not message_id or not new_status:
        return jsonify({"success": False, "error": "Missing message_id or status"}), 400
    try:
        result = db.update_status(message_id, new_status)
        return jsonify({"success": result})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/addin/stale_alerts", methods=["GET"])
def addin_stale_alerts():
    """Return emails that have not had a status change in 24+ hours."""
    try:
        stale = db.get_stale_emails(hours=24)
        return jsonify({"success": True, "stale_count": len(stale), "stale": stale})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/addin/status_flows", methods=["GET"])
def addin_status_flows():
    """Return the full status flow definitions for the frontend stepper."""
    from database import STATUS_FLOWS, NEXT_STATUS
    return jsonify({"success": True, "flows": STATUS_FLOWS, "next_status": NEXT_STATUS})

def seed_mock_data():
    session = db.Session()
    try:
        from database import ProcessedEmail, ClassificationStats
        from datetime import datetime, timedelta
        
        # Check if already seeded
        if session.query(ProcessedEmail).count() > 0:
            return
            
        print("Seeding mock data for dashboard demonstration...")
        
        # Extracted mock details
        mock_po_details = {
            "Customer": "BoxMakers Co.",
            "PO Number": "99231",
            "Quantity": "2,500 units",
            "Items Ordered": "Custom Corrugated Boxes (3-Ply)",
            "Delivery Date": "July 12, 2026"
        }
        
        mock_enquiry_details = {
            "Customer": "QuickBox Inc.",
            "Requested Item": "Heavy Duty Cartons (5-Ply)",
            "Quantity": "5,000 units",
            "Custom Printing (Yes/No)": "Yes (Logo)",
            "Samples Requested (Yes/No)": "Yes",
            "Urgency": "High"
        }
        
        mock_invoice_details = {
            "Vendor": "FastFreight Logistics",
            "Invoice Number": "INV-2026-4401",
            "Amount Due": "$3,450.00",
            "Due Date": "July 29, 2026",
            "Payment Terms": "Net 30"
        }
        
        mock_shipping_details = {
            "Carrier": "DHL Express",
            "Tracking Number": "88201882",
            "Est Delivery Date": "Wednesday, July 1",
            "Current Status": "In Transit"
        }
        
        mock_general_details = {
            "Topic": "Friday Team Lunch",
            "Action Item": "Submit preference before deadline"
        }
        
        # Mock drafts
        po_draft = "Dear Procurement Team,\n\nThank you for your purchase order PO #99231. We have received it and have scheduled it for production.\n\nOrder details:\n- Item: Custom Corrugated Boxes (3-Ply)\n- Quantity: 2,500 units\n- Target Delivery Date: July 12, 2026\n\nIf you have any questions, please let us know.\n\nBest regards,\nProvident Packaging Operations"
        
        enquiry_draft = "Hi Sales Team at QuickBox,\n\nThank you for reaching out. We have received your inquiry for 5,000 custom-printed heavy duty cartons with your logo, as well as your request for samples.\n\nOur engineering team is calculating the pricing estimate and preparing the samples. We will send you the quotation and sample dispatch tracking details within the next 24 hours.\n\nBest regards,\nProvident Packaging Sales Team"
        
        invoice_draft = "Hello Billing Team,\n\nWe have received invoice #INV-2026-4401 in the amount of $3,450.00. It has been routed to our accounts payable department for verification. It will be scheduled for payment in accordance with our Net 30 terms (due July 29, 2026).\n\nBest regards,\nProvident Packaging Accounts Payable"
        
        shipping_draft = "Hi Logistics Team,\n\nThank you for the tracking details. We have received the shipment notification (DHL Ref 88201882) and will monitor its progress for arrival on Wednesday.\n\nBest regards,\nProvident Packaging Logistics"
        
        general_draft = "Hi team,\n\nThanks for the reminder. I will make sure to submit my preference before the deadline.\n\nBest regards,\nProvident Packaging Operations"

        # Create mock processed emails
        mock_emails = [
            ProcessedEmail(
                message_id="mock-msg-1",
                user_email="info@providentpackaging.com",
                sender="procurement@boxmakers.co",
                subject="PO #99231 - Corrugated Box Order",
                body_preview="Please find attached our purchase order for 2,500 units of custom boxes. Delivery needed next week.",
                category="purchase_order",
                confidence=95.0,
                reason="Email contains clear 'purchase order' request and attachment details.",
                outlook_category_applied="Purchase Order",
                is_correct=True,
                received_at=datetime.utcnow() - timedelta(hours=2),
                extracted_data=json.dumps(mock_po_details),
                response_draft=po_draft
            ),
            ProcessedEmail(
                message_id="mock-msg-2",
                user_email="info@providentpackaging.com",
                sender="sales@quickbox.com",
                subject="Request for Quotation: Heavy Duty Cartons",
                body_preview="Can we get a price estimate for 5,000 units of 3-ply heavy duty cartons shipped to our warehouse? Please provide samples.",
                category="enquiry",
                confidence=92.0,
                reason="RFQ intent and pricing requests present.",
                outlook_category_applied="Enquiry",
                is_correct=None,
                received_at=datetime.utcnow() - timedelta(hours=4),
                extracted_data=json.dumps(mock_enquiry_details),
                response_draft=enquiry_draft
            ),
            ProcessedEmail(
                message_id="mock-msg-3",
                user_email="logistics@providentpackaging.com",
                sender="billing@fastfreight.com",
                subject="Invoice INV-2026-4401 due for payment",
                body_preview="Your invoice for last week's shipping logistics is now available. Amount due: $3,450. Please settle within net 30.",
                category="invoice",
                confidence=98.0,
                reason="Explicit invoice number and amount due mentioned in body.",
                outlook_category_applied="Invoice",
                is_correct=None,
                received_at=datetime.utcnow() - timedelta(hours=6),
                extracted_data=json.dumps(mock_invoice_details),
                response_draft=invoice_draft
            ),
            ProcessedEmail(
                message_id="mock-msg-4",
                user_email="ops@providentpackaging.com",
                sender="tracking@dhl.com",
                subject="DHL Shipment Dispatch Notification: Ref 88201",
                body_preview="Your shipment with waybill number 88201882 has been dispatched from our terminal and is in transit. Est delivery: Wed.",
                category="shipping",
                confidence=94.0,
                reason="Logistics tracking number and dispatch status.",
                outlook_category_applied="Shipping",
                is_correct=True,
                received_at=datetime.utcnow() - timedelta(days=1),
                extracted_data=json.dumps(mock_shipping_details),
                response_draft=shipping_draft
            ),
            ProcessedEmail(
                message_id="mock-msg-5",
                user_email="info@providentpackaging.com",
                sender="team-lunch@provident.com",
                subject="Friday Team Lunch",
                body_preview="Hi team, we are ordering lunch this Friday. Please submit your preferences by tomorrow noon. Thanks!",
                category="general",
                confidence=85.0,
                reason="General correspondence regarding internal lunch plans.",
                outlook_category_applied="General",
                is_correct=None,
                received_at=datetime.utcnow() - timedelta(days=1, hours=3),
                extracted_data=json.dumps(mock_general_details),
                response_draft=general_draft
            )
        ]
        
        session.add_all(mock_emails)
        
        # Create daily stats for the last 7 days
        today = datetime.utcnow()
        for i in range(7):
            date_str = (today - timedelta(days=i)).strftime("%Y-%m-%d")
            # Create stats for a couple of categories
            session.add(ClassificationStats(
                date=date_str,
                user_email="info@providentpackaging.com",
                category="purchase_order",
                count=3 + (i % 3),
                avg_confidence=92.5
            ))
            session.add(ClassificationStats(
                date=date_str,
                user_email="info@providentpackaging.com",
                category="enquiry",
                count=2 + (i % 2),
                avg_confidence=89.0
            ))
            session.add(ClassificationStats(
                date=date_str,
                user_email="info@providentpackaging.com",
                category="invoice",
                count=1 + (i % 2),
                avg_confidence=96.0
            ))
            
        session.commit()
        print("Mock data seeded successfully.")
    except Exception as e:
        session.rollback()
        print(f"Error seeding mock data: {e}")
    finally:
        session.close()

if __name__ == "__main__":
    print("=" * 60)
    print("PROVIDENT EMAIL CLASSIFIER - WEB DASHBOARD & ADD-IN")
    if "--ssl" in sys.argv:
        print("Starting secure server on: https://localhost:7071")
    else:
        print("Starting server on: http://localhost:7071")
    print("=" * 60)
    seed_mock_data()
    
    if "--ssl" in sys.argv:
        import os as _os
        _base = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
        _cert = _os.path.join(_base, "certs", "cert.pem")
        _key  = _os.path.join(_base, "certs", "key.pem")
        if _os.path.exists(_cert) and _os.path.exists(_key):
            print(f"Using permanent SSL cert: {_cert}")
            app.run(host="0.0.0.0", port=7071, debug=True, ssl_context=(_cert, _key))
        else:
            print("WARNING: certs/cert.pem not found — falling back to adhoc SSL")
            print("Run: openssl req -x509 -newkey rsa:2048 -keyout certs/key.pem -out certs/cert.pem -days 365 -nodes -subj '/CN=localhost' -addext 'subjectAltName=IP:127.0.0.1,DNS:localhost'")
            app.run(host="0.0.0.0", port=7071, debug=True, ssl_context="adhoc")
    else:
        app.run(host="0.0.0.0", port=7071, debug=True)

