# SSN3PCF
cat > README.md <<'EOF'
# Finnish SSN PCF Control 🇫🇮

Custom Power Apps Component Framework (PCF) control for handling Finnish personal identity codes (HETU / SSN).

This control provides:
- ✅ Format validation (DDMMYY-123A)
- ✅ Date validation (real calendar dates)
- ✅ Checksum validation
- ✅ Duplicate detection (Dataverse)
- ✅ Secure masking in read mode
- ✅ Temporary reveal (3 seconds)
- ✅ Clean, modern UI

---

## ✨ Features

### Validation
- Validates structure: `DDMMYYCZZZQ`
- Validates date correctness
- Validates checksum using official algorithm

### Duplicate detection
- Checks if SSN already exists in Dataverse
- Ignores current record (update scenario)
- Debounced + safe (no hanging requests)

### Security
- Masks SSN in read mode: `170446-***M`
- Reveal button shows full value for 3 seconds only

### UX
- 🇫🇮 Finnish flag + SSN label
- Inline validation feedback
- Spinner during duplicate check
- Clean field styling (Fluent-like)

---

## 🧱 Tech Stack

- PCF (Power Apps Component Framework)
- TypeScript
- Webpack (via `pcf-scripts`)
- Dataverse WebAPI

---

## 📦 Project structure
