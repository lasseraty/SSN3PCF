import { IInputs, IOutputs } from "./generated/ManifestTypes";

type StatusState = "empty" | "checking" | "valid" | "invalid" | "duplicate";

export class FinnishSSNControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container!: HTMLDivElement;
    private _root!: HTMLDivElement;
    private _field!: HTMLDivElement;
    private _input!: HTMLInputElement;
    private _status!: HTMLDivElement;
    private _error!: HTMLDivElement;

    private _context!: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged!: () => void;

    private _value: string = "";
    private _debounceTimer: number | null = null;
    private _requestId = 0;
    private _statusState: StatusState = "empty";

    private readonly SSN_REGEX = /^(\d{2})(\d{2})(\d{2})([-+A])(\d{3})([0-9A-FHJKLMNPRSTUVWXY])$/i;
    private readonly CHECKSUM_CHARS = "0123456789ABCDEFHJKLMNPRSTUVWXY";

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        this._value = context.parameters.ssnValue.raw || "";

        this._root = document.createElement("div");
        this._root.className = "ssn-root";

        this._field = document.createElement("div");
        this._field.className = "ssn-container";

        const left = document.createElement("div");
        left.className = "ssn-left";

        const flag = document.createElement("div");
        flag.className = "ssn-flag";
        flag.textContent = "FI";

        const label = document.createElement("div");
        label.className = "ssn-label";
        label.textContent = "SSN";

        left.appendChild(flag);
        left.appendChild(label);

        this._input = document.createElement("input");
        this._input.className = "ssn-input";
        this._input.type = "text";
        this._input.placeholder = "DDMMYY-123A";
        this._input.maxLength = 11;
        this._input.value = this._value;
        this._input.setAttribute("autocomplete", "off");
        this._input.setAttribute("spellcheck", "false");

        if (this._context.mode.isControlDisabled) {
            this._input.disabled = true;
        }

        this._input.addEventListener("input", this._onInput.bind(this));
        this._input.addEventListener("blur", this._onBlur.bind(this));

        this._status = document.createElement("div");
        this._status.className = "ssn-status";
        this._status.innerHTML = this._svgEmpty();

        this._error = document.createElement("div");
        this._error.className = "ssn-error-text";

        this._field.appendChild(left);
        this._field.appendChild(this._input);
        this._field.appendChild(this._status);

        this._root.appendChild(this._field);
        this._root.appendChild(this._error);

        this._container.innerHTML = "";
        this._container.appendChild(this._root);

        if (this._value) {
            this._runValidationFlow(this._value);
        } else {
            this._setStatus("empty");
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        const newValue = context.parameters.ssnValue.raw || "";

        if (newValue !== this._value) {
            this._value = newValue;
            if (this._input) {
                this._input.value = newValue;
            }

            if (newValue) {
                this._runValidationFlow(newValue);
            } else {
                this._setStatus("empty");
                this._setError("");
            }
        }

        if (this._input) {
            this._input.disabled = this._context.mode.isControlDisabled;
        }
    }

    public getOutputs(): IOutputs {
        return {
            ssnValue: this._value
        };
    }

    public destroy(): void {
        if (this._debounceTimer) {
            clearTimeout(this._debounceTimer);
            this._debounceTimer = null;
        }
    }

    private _onInput(): void {
        const raw = this._input.value;
        const formatted = this._normalizeInput(raw);

        if (formatted !== raw) {
            const pos = this._input.selectionStart || formatted.length;
            this._input.value = formatted;
            this._input.setSelectionRange(pos, pos);
        }

        this._value = formatted;
        this._notifyOutputChanged();

        if (!formatted) {
            this._requestId++;
            this._setStatus("empty");
            this._setError("");
            return;
        }

        this._runValidationFlow(formatted);
    }

    private _onBlur(): void {
        if (!this._value) return;
        this._runValidationFlow(this._value, true);
    }

    private _runValidationFlow(value: string, immediate: boolean = false): void {
        const requestId = ++this._requestId;

        if (this._debounceTimer) {
            clearTimeout(this._debounceTimer);
            this._debounceTimer = null;
        }

        const formatCheck = this._validateSSN(value);

        if (!formatCheck.valid) {
            this._setStatus("invalid");
            this._setError(formatCheck.message || "Invalid SSN");
            return;
        }

        this._setStatus("checking");
        this._setError("");

        const doCheck = () => {
            this._checkDuplicate(value, requestId);
        };

        if (immediate) {
            doCheck();
        } else {
            this._debounceTimer = window.setTimeout(doCheck, 450);
        }
    }

    private _normalizeInput(value: string): string {
        return value.replace(/\s+/g, "").toUpperCase();
    }

    private _validateSSN(value: string): { valid: boolean; message?: string } {
        const match = value.match(this.SSN_REGEX);

        if (!match) {
            return { valid: false, message: "Use format DDMMYY-123A" };
        }

        const dd = parseInt(match[1], 10);
        const mm = parseInt(match[2], 10);
        const yy = parseInt(match[3], 10);
        const centuryMarker = match[4];
        const individual = match[5];
        const checksum = match[6].toUpperCase();

        const century =
            centuryMarker === "+" ? 1800 :
            centuryMarker === "-" ? 1900 :
            2000;

        const fullYear = century + yy;
        const date = new Date(fullYear, mm - 1, dd);

        if (
            date.getFullYear() !== fullYear ||
            date.getMonth() !== mm - 1 ||
            date.getDate() !== dd
        ) {
            return { valid: false, message: "Invalid date in SSN" };
        }

        const controlNumber = `${match[1]}${match[2]}${match[3]}${individual}`;
        const expectedChecksum = this.CHECKSUM_CHARS[parseInt(controlNumber, 10) % 31];

        if (checksum !== expectedChecksum) {
            return { valid: false, message: `Invalid checksum, expected ${expectedChecksum}` };
        }

        return { valid: true };
    }

    private _checkDuplicate(value: string, requestId: number): void {
        const entityName =
            (this._context as any).page?.entityTypeName ||
            (this._context as any).page?.getEntityName?.() ||
            "";

        let attributeName = "";
        try {
            attributeName =
                (this._context.parameters.ssnValue as any)?.attributes?.LogicalName ||
                (this._context.parameters.ssnValue as any)?._property?.name ||
                "";
        } catch (_) {}

        if (!entityName || !attributeName) {
            if (requestId !== this._requestId) return;
            this._setStatus("valid");
            this._setError("");
            return;
        }

        const currentIdRaw: string = (this._context as any).page?.entityId || "";
        const currentId = currentIdRaw.replace(/[{}]/g, "");
        const escapedValue = value.replace(/'/g, "''");

        const selfFilter = currentId ? ` and ${entityName}id ne ${currentId}` : "";
        const query = `?$select=${attributeName}&$filter=${attributeName} eq '${escapedValue}'${selfFilter}&$top=1`;

        const timeoutPromise = new Promise<never>((_, reject) => {
            window.setTimeout(() => reject(new Error("Duplicate check timeout")), 4000);
        });

        Promise.race([
            this._context.webAPI.retrieveMultipleRecords(entityName, query),
            timeoutPromise
        ])
            .then((result: any) => {
                if (requestId !== this._requestId) return;

                if (result?.entities?.length > 0) {
                    this._setStatus("duplicate");
                    this._setError("Duplicate SSN found");
                    return;
                }

                this._setStatus("valid");
                this._setError("");
            })
            .catch((err: any) => {
                if (requestId !== this._requestId) return;

                this._setStatus("valid");
                this._setError("");

                // eslint-disable-next-line no-console
                console.warn("SSN duplicate check failed:", err);
            });
    }

    private _setStatus(state: StatusState): void {
        this._statusState = state;
        this._field.className = "ssn-container";

        switch (state) {
            case "empty":
                this._status.innerHTML = this._svgEmpty();
                break;
            case "checking":
                this._status.innerHTML = this._svgSpinner();
                break;
            case "valid":
                this._field.className += " ssn-valid";
                this._status.innerHTML = this._svgTick();
                break;
            case "invalid":
                this._field.className += " ssn-invalid";
                this._status.innerHTML = this._svgCross();
                break;
            case "duplicate":
                this._field.className += " ssn-invalid";
                this._status.innerHTML = this._svgCross();
                break;
        }
    }

    private _setError(message: string): void {
        this._error.textContent = message;
    }

    private _svgEmpty(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
          <circle cx="8" cy="8" r="6.5" stroke="#C8C6C4" />
        </svg>`;
    }

    private _svgSpinner(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true" class="ssn-spinner">
          <circle cx="8" cy="8" r="6" stroke="#C8C6C4" stroke-width="1.5" stroke-dasharray="10 8" stroke-linecap="round" />
        </svg>`;
    }

    private _svgTick(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
          <circle cx="8" cy="8" r="7" fill="#DCFCE7" stroke="#86EFAC" stroke-width="1"/>
          <polyline points="5,8.5 7,10.5 11,6" stroke="#16A34A" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
        </svg>`;
    }

    private _svgCross(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
          <circle cx="8" cy="8" r="7" fill="#FEE2E2" stroke="#FCA5A5" stroke-width="1"/>
          <line x1="5.5" y1="5.5" x2="10.5" y2="10.5" stroke="#DC2626" stroke-width="1.5" stroke-linecap="round"/>
          <line x1="10.5" y1="5.5" x2="5.5" y2="10.5" stroke="#DC2626" stroke-width="1.5" stroke-linecap="round"/>
        </svg>`;
    }
}
