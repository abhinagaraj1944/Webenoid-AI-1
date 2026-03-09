# Webenoid AI Backend Documentation

## Introduction
The Webenoid AI Backend is a robust FastAPI-based application designed to process spreadsheet data using natural language processing (NLP). It powers the Webenoid Excel Add-in, providing capabilities such as data analysis, chart generation, and dynamic dashboards.

## Technology Stack
- **Framework**: [FastAPI](https://fastapi.tiangolo.com/)
- **Database**: PostgreSQL (via SQLAlchemy ORM)
- **Data Processing**: Pandas, NumPy
- **Authentication**: Bcrypt hashing
- **Integration**: OpenAI/LLM-based engines for data interpretation

---

## API Endpoints

### 1. General Endpoints

#### **GET /**
- **Description**: Health check endpoint to verify if the backend is running.
- **Response**: `{"message": "Webenoid AI Running 🚀"}`

#### **GET /icon/{filename}**
- **Description**: Serves icon files for the Excel add-in.
- **Path Parameters**: `filename` (string)
- **Response**: Image file or error message.

---

### 2. Authentication Endpoints (`/auth`)

#### **POST /auth/signup**
- **Description**: Registers a new user.
- **Request Body**:
  ```json
  {
    "name": "Full Name",
    "email": "user@example.com",
    "phone": "1234567890",
    "password": "securepassword"
  }
  ```
- **Validation**:
  - Validates email format.
  - Validates phone format (10-15 digits).
  - Checks for existing email in the database.
- **Response**: Success message and user details.

#### **POST /auth/login**
- **Description**: Authenticates a user and returns their profile.
- **Request Body**:
  ```json
  {
    "email": "user@example.com",
    "password": "securepassword"
  }
  ```
- **Response**: Success message and user details.

---

### 3. Data Processing Endpoints

#### **POST /analyze**
- **Description**: Analyzes local Excel data based on a natural language question.
- **Request Body**:
  ```json
  {
    "question": "What is the total sales for March?",
    "data": { "Sheet1": [ ...records... ] },
    "user_name": "Abhishek",
    "user_email": "abhishek@example.com"
  }
  ```
- **Processing Logic**:
  - Handles greetings and common help queries.
  - Uses `ExcelAgent` to interpret the question and process the data.
  - Saves the query history into the database (question, response, chart type, time, user info).
- **Response**: Analysis result (text, chart data, or table data).

#### **POST /dashboard**
- **Description**: Generates a dynamic dashboard dataset from the provided spreadsheet data.
- **Request Body**: Same as `/analyze` (usually triggered without a specific question).
- **Processing Logic**:
  - Combines multiple sheets.
  - Automatically detects Date, Numeric, and Categorical columns.
  - Uses `DashboardEngine` to build a structured dashboard representation.
- **Response**: `{"success": true, "dashboard": { ... } }`

---

## Database Architecture

The system uses two primary tables:

### 1. `users`
- Stores user credentials and profile information.
- Fields: `id`, `name`, `email` (unique), `phone`, `hashed_password`, `created_at`.

### 2. `query_history`
- Logs all analysis requests for auditing and user history.
- Fields: `id`, `question`, `ai_response`, `chart_type`, `created_at`, `user_name`, `user_email`, `response_type`.

---

## Core Components

The backend logic is modularized into several "engines," each responsible for a specific part of the data processing pipeline:

### Primary Engines
- **`ExcelAgent`**: The main interface that coordinates between the user's question and the various underlying engines.
- **`QueryEngine`**: Orchestrates the natural language to data operation translation.
- **`DashboardEngine`**: Specifically designed to build structured data for the dynamic dashboard view, including multi-chart configurations.

### Specialized Engines (in `backend/engines/`)
- **`AI Engine`**: Manages communication with LLM services (e.g., OpenAI).
- **`Intent Engine`**: Determines the user's goal (e.g., "counting," "charting," "summarizing").
- **`Python Engine`**: Dynamically generates and executes Python code for complex data manipulations.
- **`Aggregation Engine`**: Performs mathematical operations like sums, averages, and counts.
- **`Condition Engine`**: Handles filtering logic and mapping natural language conditions to data filters.
- **`Column Profiler`**: Analyzes spreadsheet columns to detect data types, categories, and potential metrics.
- **`Data Cleaner`**: Robustly handles formatting issues and missing values in Excel data.
- **`Insight Engine`**: Generates human-readable summaries and "smart insights" from raw results.
- **`Schema Engine`**: Understands the structure and relationships within the spreadsheet data.
- **`Memory Engine`**: Provides context-aware processing for multi-turn conversations (if enabled).

## Setup and Installation

1. **Environment Variables**: Create a `.env` file in the `backend/` directory with:
   ```env
   DATABASE_URL=postgresql://user:password@localhost/dbname
   ```
2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Initialize Database**:
   ```bash
   python database/database.py
   ```
4. **Run the Server**:
   ```bash
   uvicorn main:app --reload
   ```
