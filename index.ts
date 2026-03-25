import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class FinnishSSNControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container!: HTMLDivElement;
    private _root!: HTMLDivElement;
    private _input!: HTMLInputElement;
    private _status!: HTMLDivElement;
    private _error!: HTMLDivElement;
    private _value: string = "";

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._container = container;
        this._value = context.parameters.ssnValue.raw || "";

        this._root = document.createElement("div");
        this._root.className = "ssn-root";

        const field = document.createElement("div");
        field.className = "ssn-container";

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

        this._input.addEventListener("input", () => {
            this._value = this._input.value;
            this._clearError();
            notifyOutputChanged();
        });

        this._status = document.createElement("div");
        this._status.className = "ssn-status";
        this._status.innerHTML = this._svgEmpty();

        this._error = document.createElement("div");
        this._error.className = "ssn-error-text";

        field.appendChild(left);
        field.appendChild(this._input);
        field.appendChild(this._status);

        this._root.appendChild(field);
        this._root.appendChild(this._error);

        this._container.innerHTML = "";
        this._container.appendChild(this._root);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const newValue = context.parameters.ssnValue.raw || "";
        if (newValue !== this._value) {
            this._value = newValue;
            if (this._input) {
                this._input.value = newValue;
            }
        }
    }

    public getOutputs(): IOutputs {
        return {
            ssnValue: this._value
        };
    }

    public destroy(): void {}

    private _clearError(): void {
        if (this._error) {
            this._error.textContent = "";
        }
    }

    private _svgEmpty(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
          <circle cx="8" cy="8" r="6.5" stroke="#C8C6C4" />
        </svg>`;
    }
}
