# Intelligent Resource Allocation in Cloud Computing and Storage System

**A Machine Learning-Driven Telemetry & Dynamic Elastic Scaling Platform**  
*(Academic Project / Final Year Submission)*

---

## 📌 Executive Summary
In multi-tenant cloud data centers, static over-provisioning leads to severe financial waste, while under-provisioning causes SLA violations and system outages. This project presents an end-to-end intelligent cloud management system combining:
1. **Live Host Telemetry:** Real-time tracking of CPU, RAM, and Disk Storage via continuous hardware polling (`psutil`).
2. **Machine Learning Allocation Engine:** A trained **Random Forest Classifier (96.5% Accuracy)** that predicts optimal resource tiers (`LOW`, `MEDIUM`, `HIGH`), provisioned VM clusters, and estimated hourly operating costs.
3. **Interactive Control Dashboard:** A dark-themed Flask-based management console with real-time multi-series HTML5 Canvas graphs and an interactive workload simulation testbed.
4. **Research & Modeling Notebook (`Cloud_Resource_Allocation_ML.ipynb`):** Complete ML development lifecycle including synthetic workload synthesis, exploratory data analysis, algorithm comparison, confusion matrix analysis, and feature importance weighting.

---

## 🏗️ System Architecture

```text
+-------------------------------------------------------------+
|                  Cloud Management Dashboard                  |
|     (Dark UI, HTML5 Canvas Live Chart, Telemetry Stream)     |
+------------------------------+------------------------------+
                               | (REST / WebSocket-style Polling)
+------------------------------v------------------------------+
|                     Flask Backend (app.py)                   |
|  - Real-time Hardware Telemetry (psutil)                    |
|  - Request Routing & Simulator Endpoint                      |
+------------------------------+------------------------------+
                               |
               +---------------+---------------+
               |                               |
+--------------v---------------+ +-------------v--------------+
|  ML Inference Engine         | |     Persistent Storage      |
|  (Random Forest: .joblib)    | |  (SQLite: allocation DB)    |
|  - Input: CPU, RAM, Disk,    | |  - Telemetry snapshots      |
|    Concurrent Tasks, I/O     | |  - Audit logs               |
|  - Output: Scaled VM Cluster | |  - CSV Export Engine        |
+------------------------------+ +-----------------------------+
```

---

## 🚀 Key Features

- **Multi-Resource Real-time Monitoring:** Tracks CPU %, RAM %, and Storage used/total GB.
- **Predictive Demand Assessment:** Evaluates workload pressure and automates horizontal/vertical scaling decisions.
- **Cost Estimation Matrix:**
  - `LOW Demand`: 1x `t3.micro` VM (\$0.024/hr)
  - `MEDIUM Demand`: 2x `t3.medium` VM (\$0.084/hr)
  - `HIGH Demand`: 4x `t3.xlarge` VM + Scaled SSD Storage (\$0.248/hr)
- **High Resource Warning Banner:** Triggers critical warnings if any metric exceeds 80%.
- **One-Click Audit Export:** Export complete historical telemetry and scaling actions as CSV for compliance.

---

## 📊 Machine Learning Model Benchmarks

| Algorithm | Test Accuracy | Precision (Weighted) | Recall (Weighted) | Status |
| :--- | :--- | :--- | :--- | :--- |
| **Random Forest Classifier** | **96.50%** | **0.97** | **0.96** | **Selected (Production)** |
| Decision Tree Classifier | 92.30% | 0.92 | 0.92 | Evaluated |
| Logistic Regression | 84.80% | 0.85 | 0.85 | Baseline |

### Feature Importance Weights:
- **CPU Demand:** ~36.8%
- **RAM Demand:** ~28.4%
- **Concurrent Tasks:** ~18.2%
- **Storage Volume Demand:** ~11.6%
- **Network I/O Throughput:** ~5.0%

---

## 📁 Project Structure

```text
Intelligence reasource allocation in cloude computting and storege system/
├── app.py                                # Main Flask Application with ML Engine
├── cloud_resource_allocator_model.joblib # Trained Random Forest Model
├── cloud_workload_dataset.csv            # 3000-sample Cloud Workload Dataset
├── Cloud_Resource_Allocation_ML.ipynb    # Academic Research Notebook (All steps)
├── Untitled-1.ipynb                      # Jupyter Notebook for VS Code
├── resource_allocation.db                # SQLite database for persistent logging
├── run.bat                               # Windows 1-click execution batch script
├── requirements.txt                      # Python dependencies list
├── dataset_eda_plot.png                  # Generated EDA visualization chart
├── confusion_matrix.png                  # Generated model confusion matrix
├── feature_importance.png                # Generated feature importance chart
├── README.md                             # Quick project overview
└── PROJECT_REPORT.md                     # Academic submission & viva guide
```

---

## 💻 How to Run the Project

### 1. Launch the Web Application
Double-click `run.bat` or open your terminal and run:
```bash
python app.py
```
Open your browser and navigate to:
👉 **http://127.0.0.1:5000**

### 2. View/Run the Jupyter Notebook
Open `Cloud_Resource_Allocation_ML.ipynb` or `Untitled-1.ipynb` in VS Code or Jupyter Notebook:
```bash
jupyter notebook Cloud_Resource_Allocation_ML.ipynb
```
