/** Errors with a message that is safe and useful to show in the app. */
export class AIError extends Error {
  constructor(
    message: string,
    readonly status: number,
  ) {
    super(message);
  }
}
