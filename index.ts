import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class FinnishSSNControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _currentValue: string;
    private _wrapper: HTMLDivElement;
    private _input: HTMLInputElement;
    private _statusIcon: HTMLDivElement;
    private _lockedValue: HTMLDivElement;
    private _eyeBtn: HTMLButtonElement;
    private _hintEl: HTMLDivElement;
    private _labelEl: HTMLDivElement;
    private _debounceTimer: number | null = null;
    private _revealTimer: number | null = null;
    private _countdownTimer: number | null = null;
    private _isRevealed: boolean = false;
    private _duplicateRequestId: number = 0;
    private _lastValidatedValue: string = "";

    private readonly SSN_REGEX = /^(\d{2})(\d{2})(\d{2})([-+A])(\d{3})([0-9A-FHJKLMNPRSTUVWXY])$/i;
    private readonly CHECKSUM_CHARS = "0123456789ABCDEFHJKLMNPRSTUVWXY";

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        this._currentValue = context.parameters.ssnValue.raw || "";
        this._render();
    }

    private _isEditMode(): boolean {
        return !this._context.mode.isControlDisabled;
    }

    private _render(): void {
        this._container.innerHTML = "";

        const card = document.createElement("div");
        card.className = "ssn-card";

        const header = document.createElement("div");
        header.className = "ssn-card-header";

        this._labelEl = document.createElement("div");
        this._labelEl.className = "ssn-label";
        this._labelEl.innerHTML = `<span class="ssn-flag">🇫🇮</span><span>SSN</span>`;

        header.appendChild(this._labelEl);
        card.appendChild(header);

        this._wrapper = document.createElement("div");
        this._wrapper.className = "ssn-field-row";

        if (this._isEditMode()) {
            this._renderEdit();
        } else {
            this._renderLocked();
        }

        card.appendChild(this._wrapper);

        this._hintEl = document.createElement("div");
        this._hintEl.className = "ssn-hint";
        card.appendChild(this._hintEl);

        this._container.appendChild(card);

        this._setHint(
            this._isEditMode()
                ? "Enter Finnish personal identity code"
                : "Click eye to reveal for 3 seconds",
            ""
        );
    }

    private _renderEdit(): void {
        this._input = document.createElement("input");
        this._input.type = "text";
        this._input.className = "ssn-input";
        this._input.maxLength = 11;
        this._input.placeholder = "DDMMYY-123A";
        this._input.value = this._currentValue;
        this._input.setAttribute("autocomplete", "off");
        this._input.addEventListener("input", this._onInput.bind(this));

        this._statusIcon = document.createElement("div");
        this._statusIcon.className = "ssn-status";
        this._statusIcon.innerHTML = this._svgEmpty();

        this._wrapper.appendChild(this._input);
        this._wrapper.appendChild(this._statusIcon);
    }

    private _renderLocked(): void {
        this._lockedValue = document.createElement("div");
        this._lockedValue.className = "ssn-locked-value";
        this._lockedValue.textContent = this._mask(this._currentValue);

        this._eyeBtn = document.createElement("button");
        this._eyeBtn.className = "ssn-eye-btn";
        this._eyeBtn.title = "Reveal SSN";
        this._eyeBtn.innerHTML = this._svgEye();
        this._eyeBtn.addEventListener("click", this._onEyeClick.bind(this));

        this._wrapper.appendChild(this._lockedValue);
        this._wrapper.appendChild(this._eyeBtn);
    }

    private _onInput(): void {
        const raw = this._input.value;
        const val = raw.replace(/[a-z]/g, (c) => c.toUpperCase());

        if (val !== raw) {
            const pos = this._input.selectionStart || 0;
            this._input.value = val;
            this._input.setSelectionRange(pos, pos);
        }

        this._currentValue = val;

        if (this._debounceTimer) {
            clearTimeout(this._debounceTimer);
            this._debounceTimer = null;
        }

        this._duplicateRequestId++;

        if (!val) {
            this._setStatus("empty");
            this._setHint("Enter Finnish personal identity code", "");
            this._notifyOutputChanged();
            return;
        }

        const check = this._validateFormat(val);
        if (!check.valid) {
            this._setStatus("invalid");
            this._setHint(check.message || "Invalid SSN", "error");
            return;
        }

        this._setStatus("checking");
        this._setHint("Checking for duplicates…", "");

        const requestId = this._duplicateRequestId;
        this._debounceTimer = window.setTimeout(() => {
            this._checkDuplicate(val, requestId);
        }, 400);
    }

    private _validateFormat(ssn: string): { valid: boolean; message?: string } {
        const m = ssn.match(this.SSN_REGEX);
        if (!m) return { valid: false, message: "Invalid format — expected DDMMYY-123A" };

        const dd = parseInt(m[1], 10);
        const mm = parseInt(m[2], 10);
        const yy = parseInt(m[3], 10);

        if (mm < 1 || mm > 12) return { valid: false, message: "Invalid month" };
        if (dd < 1 || dd > 31) return { valid: false, message: "Invalid day" };

        const century = m[4] === "+" ? 1800 : m[4] === "-" ? 1900 : 2000;
        const year = century + yy;
        const date = new Date(`${year}-${m[2]}-${m[1]}`);

        if (
            isNaN(date.getTime()) ||
            date.getFullYear() !== year ||
            date.getMonth() + 1 !== mm ||
            date.getDate() !== dd
        ) {
            return { valid: false, message: "Invalid date" };
        }

        const num = parseInt(m[1] + m[2] + m[3] + m[5], 10);
        const expected = this.CHECKSUM_CHARS[num % 31];

        if (m[6].toUpperCase() !== expected) {
            return { valid: false, message: `Invalid checksum — expected '${expected}'` };
        }

        return { valid: true };
    }

    private _checkDuplicate(ssn: string, requestId: number): void {
        const entityName =
            (this._context as any).page?.entityTypeName ||
            (this._context as any).page?.getEntityName?.() ||
            "";

        let attrName = "";
        try {
            attrName =
                (this._context.parameters.ssnValue as any)?.attributes?.LogicalName ||
                (this._context.parameters.ssnValue as any)?._property?.name ||
                "";
        } catch (_) {}

        if (!entityName || !attrName) {
            if (requestId !== this._duplicateRequestId) return;
            this._lastValidatedValue = ssn;
            this._setStatus("valid");
            this._setHint("Valid format. Duplicate check unavailable in this host.", "success");
            this._notifyOutputChanged();
            return;
        }

        const currentIdRaw: string = (this._context as any).page?.entityId || "";
        const currentId = currentIdRaw.replace(/[{}]/g, "");
        const escapedSsn = ssn.replace(/'/g, "''");

        const selfFilter = currentId ? ` and ${entityName}id ne ${currentId}` : "";
        const query = `?$select=${attrName}&$filter=${attrName} eq '${escapedSsn}'${selfFilter}&$top=1`;

        const timeoutPromise = new Promise<never>((_, reject) => {
            window.setTimeout(() => reject(new Error("Duplicate check timeout")), 4000);
        });

        Promise.race([
            this._context.webAPI.retrieveMultipleRecords(entityName, query),
            timeoutPromise
        ])
            .then((result: any) => {
                if (requestId !== this._duplicateRequestId) return;

                if (result.entities && result.entities.length > 0) {
                    this._setStatus("duplicate");
                    this._setHint("Duplicate found — this SSN already exists on another record", "error");
                    return;
                }

                this._lastValidatedValue = ssn;
                this._setStatus("valid");
                this._setHint("Valid — no duplicates found", "success");
                this._notifyOutputChanged();
            })
            .catch((err: any) => {
                if (requestId !== this._duplicateRequestId) return;

                this._lastValidatedValue = ssn;
                this._setStatus("valid");
                this._setHint("Valid format — duplicate check could not be completed", "success");
                this._notifyOutputChanged();

                // eslint-disable-next-line no-console
                console.warn("SSN duplicate check failed:", err);
            });
    }

    private _onEyeClick(): void {
        if (this._revealTimer) {
            clearTimeout(this._revealTimer);
            if (this._countdownTimer) clearInterval(this._countdownTimer);
        }

        if (this._isRevealed) {
            this._conceal();
            return;
        }

        this._isRevealed = true;
        this._lockedValue.textContent = this._currentValue;

        let secs = 3;
        this._setHint(`Hiding in ${secs}s…`, "");

        this._countdownTimer = window.setInterval(() => {
            secs--;
            if (secs > 0) this._setHint(`Hiding in ${secs}s…`, "");
        }, 1000);

        this._revealTimer = window.setTimeout(() => this._conceal(), 3000);
    }

    private _conceal(): void {
        this._isRevealed = false;
        this._lockedValue.textContent = this._mask(this._currentValue);
        this._setHint("Click eye to reveal for 3 seconds", "");

        if (this._countdownTimer) clearInterval(this._countdownTimer);
        if (this._revealTimer) clearTimeout(this._revealTimer);

        this._revealTimer = null;
        this._countdownTimer = null;
    }

    private _mask(ssn: string): string {
        if (!ssn || ssn.length < 8) return ssn;
        return `${ssn.substring(0, 7)}***${ssn[ssn.length - 1]}`;
    }

    private _setStatus(state: "empty" | "invalid" | "checking" | "valid" | "duplicate"): void {
        if (!this._statusIcon) return;

        this._wrapper.className = "ssn-field-row";

        switch (state) {
            case "empty":
                this._statusIcon.innerHTML = this._svgEmpty();
                break;
            case "invalid":
                this._wrapper.className += " ssn-error";
                this._statusIcon.innerHTML = this._svgCross();
                break;
            case "checking":
                this._statusIcon.innerHTML = this._svgSpinner();
                break;
            case "valid":
                this._wrapper.className += " ssn-success";
                this._statusIcon.innerHTML = this._svgTick();
                break;
            case "duplicate":
                this._wrapper.className += " ssn-error";
                this._statusIcon.innerHTML = this._svgCross();
                break;
        }
    }

    private _setHint(text: string, type: "" | "error" | "success"): void {
        if (!this._hintEl) return;
        this._hintEl.textContent = text;
        this._hintEl.className = "ssn-hint" + (type ? ` ssn-hint--${type}` : "");
    }

    private _svgEmpty(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" stroke="#c8c6c4" stroke-width="1"/></svg>`;
    }

    private _svgCross(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" fill="#FEE2E2" stroke="#FCA5A5" stroke-width="1"/><line x1="5.5" y1="5.5" x2="10.5" y2="10.5" stroke="#DC2626" stroke-width="1.5" stroke-linecap="round"/><line x1="10.5" y1="5.5" x2="5.5" y2="10.5" stroke="#DC2626" stroke-width="1.5" stroke-linecap="round"/></svg>`;
    }

    private _svgSpinner(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" class="ssn-spinner"><circle cx="8" cy="8" r="6" stroke="#c8c6c4" stroke-width="1.5" stroke-dasharray="10 8" stroke-linecap="round"/></svg>`;
    }

    private _svgTick(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" fill="#DCFCE7" stroke="#86EFAC" stroke-width="1"/><polyline points="5,8.5 7,10.5 11,6" stroke="#16A34A" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" fill="none"/></svg>`;
    }

    private _svgEye(): string {
        return `<svg width="26" height="26" viewBox="0 0 24 24" fill="none"><path d="M1 12C1 12 5 4 12 4C19 4 23 12 23 12C23 12 19 20 12 20C5 20 1 12 1 12Z" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/><circle cx="12" cy="12" r="3" stroke="currentColor" stroke-width="1.5"/></svg>`;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        const newVal = context.parameters.ssnValue.raw || "";

        if (newVal !== this._currentValue) {
            this._currentValue = newVal;
            this._render();
        }
    }

    public getOutputs(): IOutputs {
        return { ssnValue: this._currentValue };
    }

    public destroy(): void {
        if (this._debounceTimer) clearTimeout(this._debounceTimer);
        if (this._revealTimer) clearTimeout(this._revealTimer);
        if (this._countdownTimer) clearInterval(this._countdownTimer);
    }
}
