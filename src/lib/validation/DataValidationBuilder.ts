// Imports
import { DataValidation } from './DataValidation';
import { DataValidationCriteria } from './DataValidationCriteria';
import { Range } from '../range/Range';

// Class & methods

/**
 * DataValidationBuilder class.
 */
export class DataValidationBuilder {

  private _allowInvalid: boolean = true;
  private _helpText: string = '';
  private _type: DataValidationCriteria;
  private _values: Object[] = [];

  /**
   * Constructor
   *
   */
  constructor() {
  }

  public build() : DataValidation {
    return new DataValidation(this._allowInvalid, this._helpText, this._type, this._values);
  }

  /**
   * Sets the data-validation rule to require a date.
   *
   * @returns {DataValidationBuilder}
   */
  public requireDate() : DataValidationBuilder {
    this._type = DataValidationCriteria.DATE_IS_VALID_DATE;
    return this;
  }

  /**
   * Sets the data-validation rule to require a date after the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateAfter(date: Date) : DataValidationBuilder {
    this._type = DataValidationCriteria.DATE_AFTER;
    this._values.push(date);
    return this;
  }

  /**
   * Sets the data-validation rule to require a date before the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   */
  public requireDateBefore(date: Date) : DataValidationBuilder {
    this._type = DataValidationCriteria.DATE_BEFORE;
    this._values.push(date);
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
    this._type = DataValidationCriteria.DATE_BETWEEN;
    this._values.push(start);
    this._values.push(end);
    return this;
  }

  /**
   * Sets the data-validation rule to require a date equal to the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateEqualTo(date: Date) : DataValidationBuilder {
    this._type = DataValidationCriteria.DATE_EQUAL_TO;
    this._values.push(date);
    return this;
  }

  /**
   * Sets the data-validation rule to require a date not between the given values, inclusive of the values themselves.
   * The time fields of the Date objects are ignored; only the day, month, and year fields are used.
   *
   * @param start
   * @param end
   */
  public requireDateNotBetween(start: Date, end: Date) : DataValidationBuilder {
    this._type = DataValidationCriteria.DATE_NOT_BETWEEN;
    this._values.push(start);
    this._values.push(end);
    return this;
  }

  /**
   * Sets the data-validation rule to require a date on or after the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateOnOrAfter(date: Date) : DataValidationBuilder {
    this._type = DataValidationCriteria.DATE_ON_OR_AFTER;
    this._values.push(date);
    return this;
  }

  /**
   * Sets the data-validation rule to require a date on or before the given value.
   * The time fields of the Date object are ignored; only the day, month, and year fields are used.
   *
   * @param date
   */
  public requireDateOnOrBefore(date: Date) : DataValidationBuilder {
    this._type = DataValidationCriteria.DATE_ON_OR_BEFORE;
    this._values.push(date);
    return this;
  }

  /**
   * Sets the data-validation rule to require that the given formula evaluates to true.
   *
   * @param formula
   */
  public requireFormulaSatisfied(formula: string) : DataValidationBuilder {
    this._type = DataValidationCriteria.CUSTOM_FORMULA;
    this._values.push(formula);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number between the given values, inclusive of the values themselves.
   *
   * @param start
   * @param end
   */
  public requireNumberBetween(start: number, end: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_BETWEEN;
    this._values.push(start);
    this._values.push(end);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number equal to the given value.
   *
   * @param number
   */
  public requireNumberEqualTo(number: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_EQUAL_TO;
    this._values.push(number);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number greater than the given value.
   *
   * @param number
   */
  public requireNumberGreaterThan(number: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_GREATER_THAN;
    this._values.push(number);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number greater than or equal to the given value.
   *
   * @param number
   */
  public requireNumberGreaterThanOrEqualTo(number: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_GREATER_THAN_OR_EQUAL_TO;
    this._values.push(number);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number less than the given value.
   *
   * @param number
   */
  public requireNumberLessThan(number: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_LESS_THAN;
    this._values.push(number);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number less than or equal to the given value.
   *
   * @param number
   */
  public requireNumberLessThanOrEqualTo(number: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_LESS_THAN_OR_EQUAL_TO;
    this._values.push(number);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number not between the given values, inclusive of the values themselves.
   *
   * @param start
   * @param end
   */
  public requireNumberNotBetween(start: number, end: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_NOT_BETWEEN;
    this._values.push(start);
    this._values.push(end);
    return this;
  }

  /**
   * Sets the data-validation rule to require a number not equal to the given value.
   *
   * @param number
   */
  public requireNumberNotEqualTo(number: number) : DataValidationBuilder {
    this._type = DataValidationCriteria.NUMBER_NOT_EQUAL_TO;
    this._values.push(number);
    return this;
  }

  /**
   * Sets the data-validation rule to require that the input contains the given value.
   *
   * @param text
   */
  public requireTextContains(text: string) : DataValidationBuilder {
    this._type = DataValidationCriteria.TEXT_CONTAINS;
    this._values.push(text);
    return this;
  }

  /**
   * Sets the data-validation rule to require that the input does not contain the given value.
   *
   * @param text
   */
  public requireTextDoesNotContain(text: string) : DataValidationBuilder {
    this._type = DataValidationCriteria.TEXT_DOES_NOT_CONTAIN;
    this._values.push(text);
    return this;
  }

  /**
   * Sets the data-validation rule to require that the input is equal to the given value.
   * @param text
   */
  public requireTextEqualTo(text: string) : DataValidationBuilder {
    this._type = DataValidationCriteria.TEXT_EQUAL_TO;
    this._values.push(text);
    return this;
  }

  /**
   * Sets the data-validation rule to require that the input is in the form of an email address.
   */
  public requireTextIsEmail() : DataValidationBuilder {
    this._type = DataValidationCriteria.TEXT_IS_VALID_EMAIL;
    return this;
  }

  /**
   * Sets the data-validation rule to require that the input is in the form of a URL.
   */
  public requireTextIsUrl() : DataValidationBuilder {
    this._type = DataValidationCriteria.TEXT_IS_VALID_URL;
    return this;
  }

  /**
   * Sets the data-validation rule to require that the input is equal to one of the given values.
   *
   * Sets the data-validation rule to require that the input is equal to one of the given values, with an option to hide the dropdown menu.
   * @param values
   * @param showDropdown
   */
  public requireValueInList(values: string[], showDropdown?: boolean) : DataValidationBuilder {
    this._type = DataValidationCriteria.VALUE_IN_LIST;
    this._values.push(values);
    return this;
  }

  /**
   * Sets the data-validation rule to require that the input is equal to a value in the given range.
   *
   * Sets the data-validation rule to require that the input is equal to a value in the given range, with an option to hide the dropdown menu.
   *
   * @param range
   * @param showDropdown
   */
  public requireValueInRange(range: Range, showDropdown?: boolean) : DataValidationBuilder {
    this._type = DataValidationCriteria.VALUE_IN_RANGE;
    this._values.push(range);
    return this;
  }

  /**
   * Sets whether to show a warning when input fails data validation or whether to reject the input entirely.
   * The default for new data-validation rules is true.
   *
   * @param allowInvalidData
   */
  public setAllowInvalid(allowInvalidData: boolean) : DataValidationBuilder {
    this._allowInvalid = allowInvalidData;
    return this;
  }

  /**
   * Sets the help text shown when the user hovers over the cell on which data-validation is set.
   *
   * @param helpText
   */
  public setHelpText(helpText: string) : DataValidationBuilder {
    this._helpText = helpText;
    return this;
  }

  /**
   * Sets the data-validation rule to require criteria defined in the DataValidationCriteria enum.
   *
   * @param criteria
   * @param args
   */
  public withCriteria(criteria: DataValidationCriteria, args: Object[]) : DataValidationBuilder {
    this._type = criteria;
    this._values = args;
    return this;
  }
}
