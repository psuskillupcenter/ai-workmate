# AI Workmate: Knowledge Legacy

**AI Workmate: Knowledge Legacy** is a working prototype of an AI-assisted tacit knowledge analysis and organizational knowledge chat system.

The project uses **Google Apps Script**, **Google Sheets**, and **Gemini API** to transform staff-written experiences into structured organizational knowledge and support AI-based knowledge retrieval.

## Project Purpose

Organizations often risk losing tacit knowledge when experienced staff retire or move to other roles. This project aims to preserve and transfer that knowledge by using AI to extract insights, hidden techniques, common problems, practical solutions, warnings, and lessons learned from staff-written documents.

## Key Features

- Knowledge intake from staff-written text
- AI-powered tacit knowledge analysis using Gemini API
- Knowledge categorization
- Automatic Knowledge Card generation
- AI Workmate Knowledge Chat for asking questions from the extracted knowledge base
- Google Sheets as a lightweight knowledge database
- Google Apps Script Web App as the user interface

## Technology Stack

- Google Apps Script
- Google Sheets
- Gemini API
- Google AI Studio
- HTML
- CSS
- JavaScript
- GitHub

## System Workflow

```text
Staff-Written Experience
        ↓
Knowledge Intake
        ↓
Gemini AI Analysis
        ↓
Knowledge Categorization
        ↓
Knowledge Card Generation
        ↓
AI Workmate Knowledge Chat
        ↓
Organizational Knowledge Transfer

Google Sheet Structure

The system uses four main sheets:

RAW_KNOWLEDGE
AI_ANALYSIS
KNOWLEDGE_CARD
CHAT_LOG
Main Modules
1. Knowledge Intake

Collects staff-written documents, reflections, or interview transcripts.

2. AI Analysis

Uses Gemini API to extract tacit knowledge, key insights, hidden techniques, risks, solutions, and keywords.

3. Knowledge Cards

Transforms AI analysis results into structured knowledge cards for review and reuse.

4. AI Workmate Knowledge Chat

Allows users to ask questions based on the knowledge cards stored in the system.

Prototype Status

This project is currently a working prototype developed for organizational innovation and future expansion.

Future Development
GitHub-based version control with clasp
Audio-to-knowledge workflow
Google Docs / PDF export
NotebookLM integration
User authentication and role-based access
Enhanced search and filtering
Dashboard analytics for knowledge management
Security Note

The Gemini API key is not stored in this repository. It should be configured securely in Google Apps Script Script Properties.

```text
Update project README
