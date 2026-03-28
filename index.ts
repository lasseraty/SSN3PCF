import { IInputs, IOutputs } from "./generated/ManifestTypes";
type StatusState = "empty" | "valid" | "invalid" | "duplicate";
export class FinnishSSNControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _context!: ComponentFramework.Context<IInputs>;
  private _notifyOutputChanged!: () => void;
  private _container!: HTMLDivElement;
  private _currentValue: string = "";
  private _isNewRecord: boolean = true;
  private _requestId: number = 0;
  private _debounceTimer: number | null = null;
  private _revealTimer: number | null = null;
  private _countdownTimer: number | null = null;
  private _isRevealed: boolean = false;
  private _wrapper!: HTMLDivElement;
  private _input!: HTMLInputElement;
  private _statusIcon!: HTMLDivElement;
  private _lockedValue!: HTMLDivElement;
  private _eyeBtn!: HTMLButtonElement;
  private _hintEl!: HTMLDivElement;
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
    this._currentValue = context.parameters.ssnValue.raw || "";

    // New record = no existing value AND record hasn't been saved yet
    // entityId is empty/null on unsaved new records
    const entityId = ((context as any).page?.entityId || "").replace(/[{}]/g, "");
    this._isNewRecord = !entityId;

    this._render();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
    const newValue = context.parameters.ssnValue.raw || "";

    // Once D365 gives us an entityId the record has been saved — lock it
    const entityId = ((context as any).page?.entityId || "").replace(/[{}]/g, "");
    if (entityId) {
      this._isNewRecord = false;
    }

    if (newValue === this._currentValue) return; // no change, skip re-render
    this._currentValue = newValue;
    this._render();
  }

  public getOutputs(): IOutputs {
    return { ssnValue: this._currentValue };
  }

  public destroy(): void {
    if (this._debounceTimer) clearTimeout(this._debounceTimer);
    if (this._revealTimer) clearTimeout(this._revealTimer);
    if (this._countdownTimer) clearInterval(this._countdownTimer);
  }

  // Edit mode = new unsaved record AND field not disabled
  private _isEditMode(): boolean {
    return this._isNewRecord && !this._context.mode.isControlDisabled;
  }

  private _render(): void {
    this._container.innerHTML = "";
    this._wrapper = document.createElement("div");
    this._wrapper.className = "ssn-field-row";

    const iconWrap = document.createElement("div");
    iconWrap.className = "ssn-icon-wrap";
    iconWrap.innerHTML = this._iconSVG();
    this._wrapper.appendChild(iconWrap);

    if (this._isEditMode()) {
      this._renderEdit();
    } else {
      this._renderLocked();
    }

    this._container.appendChild(this._wrapper);

    this._hintEl = document.createElement("div");
    this._hintEl.className = "ssn-hint";
    this._container.appendChild(this._hintEl);

    this._setHint(
      this._isEditMode() ? "Enter Finnish personal identity code (PPKKVV-TTTC)" : "",
      ""
    );
  }

  private _renderEdit(): void {
    this._input = document.createElement("input");
    this._input.type = "text";
    this._input.className = "ssn-input";
    this._input.maxLength = 11;
    this._input.placeholder = "PPKKVV-TTTC";
    this._input.value = this._currentValue;
    this._input.setAttribute("autocomplete", "off");
    this._input.addEventListener("input", this._onInput.bind(this));
    this._wrapper.appendChild(this._input);

    this._statusIcon = document.createElement("div");
    this._statusIcon.className = "ssn-status";
    this._statusIcon.innerHTML = this._svgEmpty();
    this._wrapper.appendChild(this._statusIcon);
  }

  private _renderLocked(): void {
    this._lockedValue = document.createElement("div");
    this._lockedValue.className = "ssn-locked-value";
    this._lockedValue.textContent = this._isRevealed
      ? this._currentValue
      : this._mask(this._currentValue);
    this._wrapper.appendChild(this._lockedValue);

    this._eyeBtn = document.createElement("button");
    this._eyeBtn.className = "ssn-eye-btn";
    this._eyeBtn.title = "Reveal SSN for 10 seconds";
    this._eyeBtn.innerHTML = this._svgEye();
    this._eyeBtn.addEventListener("click", this._onEyeClick.bind(this));
    this._wrapper.appendChild(this._eyeBtn);
  }

  private _onInput(): void {
    const raw = this._input.value;
    const upper = raw.replace(/[a-z]/g, c => c.toUpperCase());
    if (upper !== raw) {
      const pos = this._input.selectionStart || 0;
      this._input.value = upper;
      this._input.setSelectionRange(pos, pos);
    }
    this._currentValue = upper;
    this._notifyOutputChanged();

    if (this._debounceTimer) clearTimeout(this._debounceTimer);

    if (!upper) {
      this._setStatus("empty");
      this._setHint("Enter Finnish personal identity code (PPKKVV-TTTC)", "");
      return;
    }

    const check = this._validateFormat(upper);
    if (!check.valid) {
      this._setStatus("invalid");
      this._setHint(check.message || "Invalid SSN", "error");
      return;
    }

    // Format valid — debounce then run duplicate check
    this._debounceTimer = window.setTimeout(() => {
      this._requestId++;
      this._checkDuplicate(upper, this._requestId);
    }, 500);
  }

  private _onEyeClick(): void {
    if (this._revealTimer) clearTimeout(this._revealTimer);
    if (this._countdownTimer) clearInterval(this._countdownTimer);

    this._isRevealed = true;
    this._lockedValue.textContent = this._currentValue;
    this._eyeBtn.innerHTML = this._svgEyeOff();

    let remaining = 10;
    this._eyeBtn.title = `Hide SSN (${remaining}s)`;
    this._countdownTimer = window.setInterval(() => {
      remaining--;
      this._eyeBtn.title = `Hide SSN (${remaining}s)`;
      if (remaining <= 0) this._hideSSN();
    }, 1000);

    this._revealTimer = window.setTimeout(() => this._hideSSN(), 10000);
  }

  private _hideSSN(): void {
    if (this._revealTimer) { clearTimeout(this._revealTimer); this._revealTimer = null; }
    if (this._countdownTimer) { clearInterval(this._countdownTimer); this._countdownTimer = null; }
    this._isRevealed = false;
    if (this._lockedValue) this._lockedValue.textContent = this._mask(this._currentValue);
    if (this._eyeBtn) {
      this._eyeBtn.innerHTML = this._svgEye();
      this._eyeBtn.title = "Reveal SSN for 10 seconds";
    }
  }

  private _checkDuplicate(value: string, reqId: number): void {
    const entityName = (this._context as any).page?.entityTypeName || "";
    const attributeName =
      (this._context.parameters.ssnValue as any)?.attributes?.LogicalName || "";

    if (!entityName || !attributeName) {
      // Can't check — assume valid
      this._setStatus("valid");
      this._setHint("", "");
      return;
    }

    const currentId = ((this._context as any).page?.entityId || "").replace(/[{}]/g, "");
    const escaped = value.replace(/'/g, "''");

    // Exclude current record from duplicate check (relevant when editing existing record)
    const selfFilter = currentId
      ? ` and ${entityName}id ne ${currentId}`
      : "";

    // Check both active and inactive records
    const query =
      `?$select=${attributeName}` +
      `&$filter=${attributeName} eq '${escaped}'${selfFilter}` +
      `` and (statecode eq 0 or statecode eq 1)` +
      `&$top=1`;

    const timeout = new Promise<never>((_, reject) =>
      window.setTimeout(() => reject(new Error("timeout")), 5000)
    );

    Promise.race([
      this._context.webAPI.retrieveMultipleRecords(entityName, query),
      timeout,
    ])
      .then((result: any) => {
        if (reqId !== this._requestId) return; // stale response
        if (result?.entities?.length > 0) {
          this._setStatus("duplicate");
          this._setHint("Duplicate SSN — this identity code already exists", "error");
        } else {
          this._setStatus("valid");
          this._setHint("", "");
        }
      })
      .catch(() => {
        if (reqId !== this._requestId) return;
        // On error/timeout, don't block the user — show valid
        this._setStatus("valid");
        this._setHint("", "");
      });
  }

  private _validateFormat(value: string): { valid: boolean; message?: string } {
    const m = value.match(this.SSN_REGEX);
    if (!m) return { valid: false, message: "Invalid format — expected DDMMYY-TTTC" };
    const dd = parseInt(m[1], 10);
    const mm2 = parseInt(m[2], 10);
    const yy = parseInt(m[3], 10);
    if (mm2 < 1 || mm2 > 12) return { valid: false, message: "Invalid month" };
    if (dd < 1 || dd > 31) return { valid: false, message: "Invalid day" };
    const century = m[4] === "+" ? 1800 : m[4] === "-" ? 1900 : 2000;
    const fullYear = century + yy;
    const date = new Date(`${fullYear}-${m[2]}-${m[1]}`);
    if (
      isNaN(date.getTime()) ||
      date.getFullYear() !== fullYear ||
      date.getMonth() + 1 !== mm2 ||
      date.getDate() !== dd
    ) {
      return { valid: false, message: "Invalid date" };
    }
    const num = parseInt(m[1] + m[2] + m[3] + m[5], 10);
    const expected = this.CHECKSUM_CHARS[num % 31];
    if (m[6].toUpperCase() !== expected)
      return { valid: false, message: `Invalid checksum — expected '${expected}'` };
    return { valid: true };
  }

  private _mask(value: string): string {
    if (!value || value.length < 8) return value;
    // Show DDMMYY- and mask the individual part: 010190-***A
    return value.slice(0, 7) + "***" + value.slice(-1);
  }

  private _setStatus(state: StatusState): void {
    if (!this._statusIcon) return;
    switch (state) {
      case "empty":     this._statusIcon.innerHTML = this._svgEmpty();     break;
      case "valid":     this._statusIcon.innerHTML = this._svgTick();      break;
      case "invalid":   this._statusIcon.innerHTML = this._svgCross();     break;
      case "duplicate": this._statusIcon.innerHTML = this._svgDuplicate(); break;
    }
  }

  private _setHint(msg: string, type: string): void {
    if (!this._hintEl) return;
    this._hintEl.textContent = msg;
    this._hintEl.className = "ssn-hint" + (type ? ` ssn-hint-${type}` : "");
  }

  // ── Icons ────────────────────────────────────────────────────────────────

  private _iconSVG(): string {
    return `<svg width="44" height="36" viewBox="0 0 44 36" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect width="44" height="36" rx="8" fill="#f0f9ff"/>
      <!-- FI label -->
      <text x="11" y="13" text-anchor="middle" font-size="7" font-weight="700"
            fill="#0f3d91" font-family="Segoe UI,sans-serif" letter-spacing="0.5">FI</text>
      <!-- SSN label -->
      <text x="11" y="22" text-anchor="middle" font-size="7" font-weight="700"
            fill="#0f3d91" font-family="Segoe UI,sans-serif" letter-spacing="0.5">SSN</text>
      <!-- Vertical dashed divider -->
      <line x1="22" y1="6" x2="22" y2="30" stroke="#d2d0ce" stroke-width="1" stroke-dasharray="2 2"/>
      <!-- Three dots suggesting masked data -->
      <circle cx="29" cy="18" r="2" fill="#0f3d91" opacity="0.35"/>
      <circle cx="35" cy="18" r="2" fill="#0f3d91" opacity="0.35"/>
      <circle cx="41" cy="18" r="2" fill="#0ea5e9" opacity="0.7"/>
    </svg>`;
  }

  private _svgEmpty(): string {
    return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none">
      <circle cx="8" cy="8" r="6.5" stroke="#C8C6C4"/>
    </svg>`;
  }

  private _svgTick(): string {
    return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none">
      <circle cx="8" cy="8" r="7" fill="#DCFCE7" stroke="#86EFAC" stroke-width="1"/>
      <polyline points="5,8.5 7,10.5 11,6" stroke="#16A34A" stroke-width="1.5"
                stroke-linecap="round" stroke-linejoin="round" fill="none"/>
    </svg>`;
  }

  private _svgCross(): string {
    return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none">
      <circle cx="8" cy="8" r="7" fill="#FEE2E2" stroke="#FCA5A5" stroke-width="1"/>
      <line x1="5.5" y1="5.5" x2="10.5" y2="10.5" stroke="#DC2626" stroke-width="1.5" stroke-linecap="round"/>
      <line x1="10.5" y1="5.5" x2="5.5" y2="10.5" stroke="#DC2626" stroke-width="1.5" stroke-linecap="round"/>
    </svg>`;
  }

  private _svgDuplicate(): string {
    return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none">
      <circle cx="8" cy="8" r="7" fill="#FEF9C3" stroke="#FDE047" stroke-width="1"/>
      <line x1="8" y1="5" x2="8" y2="9" stroke="#CA8A04" stroke-width="1.5" stroke-linecap="round"/>
      <circle cx="8" cy="11.5" r="0.8" fill="#CA8A04"/>
    </svg>`;
  }

  private _svgEye(): string {
    return `<svg width="18" height="18" viewBox="0 0 16 16" fill="none">
      <path d="M1.5 8C2.7 5.6 5.1 4 8 4s5.3 1.6 6.5 4c-1.2 2.4-3.6 4-6.5 4S2.7 10.4 1.5 8Z"
            stroke="#605E5C" stroke-width="1.2"/>
      <circle cx="8" cy="8" r="2.2" stroke="#605E5C" stroke-width="1.2"/>
    </svg>`;
  }

  private _svgEyeOff(): string {
    return `<svg width="18" height="18" viewBox="0 0 16 16" fill="none">
      <path d="M1.5 8C2.7 5.6 5.1 4 8 4s5.3 1.6 6.5 4c-1.2 2.4-3.6 4-6.5 4S2.7 10.4 1.5 8Z"
            stroke="#0f3d91" stroke-width="1.2"/>
      <circle cx="8" cy="8" r="2.2" stroke="#0f3d91" stroke-width="1.2"/>
      <line x1="3" y1="3" x2="13" y2="13" stroke="#0f3d91" stroke-width="1.2" stroke-linecap="round"/>
    </svg>`;
  }
}