// Imports

// Class & methods
/**
 * User class.
 */
export class User {

  private _email: string;

  /**
   * Constructor
   *
   */
  constructor (email: string) {
    this._email = email;
  }

  /**
   * Gets the user's email address, if available
   *
   * @returns {string}
   */
  public getEmail (): string {
    return this._email;
  }

}
