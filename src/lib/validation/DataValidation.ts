// Imports
import { DataValidationCriteria } from './DataValidationCriteria';

// Class & methods

/**
 * DataValidationBuilder class.
 */
export class DataValidation {

  private _allowInvalid: boolean = true;
  private _helpText: string = '';
  private _type: DataValidationCriteria;
  private _values: Object[];

  /**
   * Constructor
   */
  constructor(allowInvalid: boolean, helpText: string,
              type: DataValidationCriteria,
              values: Object[]
  ) {
    this._allowInvalid = allowInvalid;
    this._helpText = helpText;
    this._type = type;
    this._values = values;
  }

  /**
   * Returns true if the rule shows a warning when input fails data validation, or false if it rejects the input entirely.
   * The default for new data-validation rules is true.
   */
  public getAllowInvalid() : boolean {
    return this._allowInvalid;
  }

  /**
   * Gets the rule's criteria type as defined in the DataValidationCriteria enum.
   */
  public getCriteriaType() : DataValidationCriteria {
    return this._type;
  }

  /**
   * Gets an array of arguments for the rule's criteria.
   */
  public getCriteriaValues() : Object[] {
    return this._values;
  }

  /**
   * Gets the rule's help text, or null if no help text is set.
   */
  public getHelpText() : string {
    return this._helpText;
  }

}
