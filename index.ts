import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class FinnishSSNControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private container!: HTMLDivElement;
  private input!: HTMLInputElement;
  private status!: HTMLDivElement;
  private error!: HTMLDivElement;

  public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void): void {
    this.container = document.createElement("div");

    const wrapper = document.createElement("div");
    wrapper.className = "ssn-container";

    const left = document.createElement("div");
    left.className = "ssn-left";

    const flag = document.createElement("img");
    flag.className = "ssn-flag";
    flag.src = "https://upload.wikimedia.org/wikipedia/commons/b/bc/Flag_of_Finland.svg";

    const label = document.createElement("span");
    label.className = "ssn-label";
    label.textContent = "SSN";

    left.appendChild(flag);
    left.appendChild(label);

    this.input = document.createElement("input");
    this.input.className = "ssn-input";
    this.input.placeholder = "DDMMYY-123A";

    this.status = document.createElement("div");
    this.status.className = "ssn-status";

    wrapper.appendChild(left);
    wrapper.appendChild(this.input);
    wrapper.appendChild(this.status);

    this.error = document.createElement("div");
    this.error.className = "ssn-error-text";

    this.container.appendChild(wrapper);
    this.container.appendChild(this.error);

    this.input.addEventListener("input", () => {
      this.error.textContent = "";
      this.status.innerHTML = "";
      notifyOutputChanged();
    });
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    const value = context.parameters?.ssnValue?.raw || "";
    this.input.value = value;
  }

  public getOutputs(): IOutputs {
    return {
      ssnValue: this.input.value
    };
  }

  public destroy(): void {}
}
