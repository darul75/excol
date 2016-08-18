// Imports
import { Errors } from '../Error';

// Class & methods

/**
 * DataValidationBuilder class.
 */
export class DataValidationBuilder {

  private _date: boolean;
  private _dateAfter: Date;
  private _dateBefore: Date;
  private _dateBetween: Date[];

  /**
   * Constructor
   *
   */
  constructor() {
  }

  /**
   * Sets the data-validation rule to require a date.
   *
   * @returns {DataValidationBuilder}
   */
  public requireDate() : DataValidationBuilder {
    this._date = true;

    return this;
  }

  /**
   * Sets the data-validation rule to require a date after the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateAfter(date: Date) : DataValidationBuilder {
    this._dateAfter = date;
    return this;
  }

  /**
   * Sets the data-validation rule to require a date before the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   */
  public requireDateBefore(date: Date) : DataValidationBuilder {
    this._dateBefore = date;
    return this;
  }

  /**
   * Sets the data-validation rule to require a date between the given values, inclusive of the values themselves.
   * The time fields of the Date objects are ignored; only the day, month, and year fields are used.
   *
   * @param start
   * @param end
   */
  public requireDateBetween(start: Date, end: Date) : DataValidationBuilder {
    this._dateBetween.push(start);
    this._dateBetween.push(end);

    return this;
  }

  /**
   * Sets the data-validation rule to require a date equal to the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateEqualTo(date) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a date not between the given values, inclusive of the values themselves.
   * The time fields of the Date objects are ignored; only the day, month, and year fields are used.
   *
   * @param start
   * @param end
   */
  public requireDateNotBetween(start, end) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a date on or after the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateOnOrAfter(date) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a date on or before the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateOnOrBefore(date) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the given formula evaluates to true.
   *
   * @param formula
   */
  public requireFormulaSatisfied(formula) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number between the given values, inclusive of the values themselves.
   *
   * @param start
   * @param end
   */
  public requireNumberBetween(start, end) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number equal to the given value.
   *
   * @param number
   */
  public requireNumberEqualTo(number) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number greater than the given value.
   *
   * @param number
   */
  public requireNumberGreaterThan(number) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number greater than or equal to the given value.
   *
   * @param number
   */
  public requireNumberGreaterThanOrEqualTo(number) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number less than the given value.
   *
   * @param number
   */
  public requireNumberLessThan(number) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number less than or equal to the given value.
   *
   * @param number
   */
  public requireNumberLessThanOrEqualTo(number) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number not between the given values, inclusive of the values themselves.
   *
   * @param start
   * @param end
   */
  public requireNumberNotBetween(start, end) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require a number not equal to the given value.
   *
   * @param number
   */
  public requireNumberNotEqualTo(number) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the input contains the given value.
   *
   * @param text
   */
  public requireTextContains(text) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the input does not contain the given value.
   *
   * @param text
   */
  public requireTextDoesNotContain(text) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the input is equal to the given value.
   * @param text
   */
  public requireTextEqualTo(text) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the input is in the form of an email address.
   */
  public requireTextIsEmail() : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the input is in the form of a URL.
   */
  public requireTextIsUrl() : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the input is equal to one of the given values.
   *
   * Sets the data-validation rule to require that the input is equal to one of the given values, with an option to hide the dropdown menu.
   * @param values
   * @param showDropdown
   */
  public requireValueInList(values, showDropdown?: boolean) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require that the input is equal to a value in the given range.
   *
   * Sets the data-validation rule to require that the input is equal to a value in the given range, with an option to hide the dropdown menu.
   *
   * @param range
   * @param showDropdown
   */
  public requireValueInRange(range, showDropdown?: boolean) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets whether to show a warning when input fails data validation or whether to reject the input entirely.
   * The default for new data-validation rules is true.
   *
   * @param allowInvalidData
   */
  public setAllowInvalid(allowInvalidData) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the help text shown when the user hovers over the cell on which data-validation is set.
   *
   * @param helpText
   */
  public setHelpText(helpText) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Sets the data-validation rule to require criteria defined in the DataValidationCriteria enum.
   *
   * @param criteria
   * @param args
   */
  public withCriteria(criteria, args) : DataValidationBuilder {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }
}
