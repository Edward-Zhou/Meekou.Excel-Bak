/* global Office */
export class DialogInput {
  name: string;
}

export class DialogEventArg {
  message: string;
  origin: string | undefined;
  type: Office.EventType;
}
