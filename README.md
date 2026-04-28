# SecureShare Phase 3 Report

## Overview

**SecureShare** is a client-facing portal that allows employers and brokers to upload pension data directly, replacing email-based submissions. Built three years ago and rolled out at scale last year for auto-enrollment, it has been an extremely successful implementation.

Data flows one-way through the portal, routed by group pension scheme numbers and transformed directly into downstream processing systems. Current supported workflows include **new entrants** and **pension contribution schedules**.

This repository tracks the Phase 3 reporting initiative: measuring SecureShare's financial impact, operational performance, and providing executive-level insights into the platform's value.

---

## Report Audience

Executives, including up to **Declan Bollinger** and **David Harney**.

---

## Reporting Scope & Key Metrics

### Financial Metrics (Primary Focus)

| Metric | Description |
|---|---|
| Turnaround time | Time from case open to investment date |
| Investment vs. collection date | Tracking lag between fund collection and investment |
| Cost of delay | Financial exposure from processing delays (methodology TBD — Finance/Actuarial input needed) |
| Client volume | Top 122 clients manage ~€900M annually; individual schemes up to €7.5M/month |

### Operational Metrics

- Submission volumes by workflow type (new entrants, contribution schedules)
- Portal adoption rates by employer/broker
- Submission success/failure rates

---

## Data Architecture

### Usher (Third-Party System)
SecureShare data lives in metadata tables within **Usher**:

| Table | Contents |
|---|---|
| Scheme metadata | Pulled from scheme details; used for routing |
| Submission history | Record of all portal submissions |
| Employer metadata | Organisation-level data for employer linkage |
| Broker metadata | Organisation-level data for broker linkage |

### Pension Planet
Financial data (investment dates, processing dates, amounts, fund information) is accessible via **Pension Planet**.

### User Data
User/PII data is stored separately from Usher due to data protection restrictions. New reporting functionality covering user and organisation queries is expected **April/May 2026**.

---

## Project Timeline

| Milestone | Target |
|---|---|
| Initial reporting deliverable | ~3 months (from project start) |
| Full project completion | Year-long initiative |
| New user/org reporting available | April/May 2026 |
| Updated portal video guide | Next week (Darren) |

No hard deadline has been set — the 3-month target is for the first reporting output.

---

## Team & Responsibilities

| Person | Role / Action |
|---|---|
| Tony | Contact Sean re: Usher/Midas data access; review metadata tables and architecture with Kieran |
| Kieran | Provide dev environment access for data exploration |
| Lorna | Share financial exposure report; connect with Amy (Actuarial) for cost-of-delay methodology; available for Usher meetings every 2 weeks |
| Jim | Connect Tony with Phil's team for visualization standards and templates |
| Amy (Actuarial) | Input on cost-of-delay calculation method |
| Phil's team | Visualization standards and report templates |
| Darren | Producing updated SecureShare video guide |

---

## Repository Structure

```
SecureShare-Phase-3-Report/
├── README.md
├── executive-summary/        # High-level findings for Declan Bollinger / David Harney
├── financial-metrics/        # Turnaround time, cost of delay, investment tracking
├── data-architecture/        # Usher/Pension Planet schema notes and data mapping
├── operational-metrics/      # Submission volumes, adoption, workflow performance
└── next-steps/               # Action items, open questions, stakeholder outputs
```

---

## Open Questions

- Cost-of-delay calculation methodology — awaiting Finance/Actuarial input (Amy)
- Midas data access requirements — Tony to confirm with Sean
- Visualization standards — pending connection with Phil's team

---

## Contact

Tony Victor — [tony@datasphered.com](mailto:tony@datasphered.com)

---

*Datasphere Dynamics — SecureShare Phase 3 Report — 2026*
