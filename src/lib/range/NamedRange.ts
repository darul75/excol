// Imports

import { Range } from './Range'
import { Spreadsheet } from '../spreadsheet/Spreadsheet';

// Class & methods
/**
 * Spreadsheet class.
 */
export class NamedRange {

  private _name: string;
  private _range: Range;
  private _parent: Spreadsheet;
  private _id: number;

  /**
   * Constructor
   *
   */
  constructor(spreadsheet: Spreadsheet, name: string, range: Range) {
    this._parent = spreadsheet;
    this._name = name;
    this._range = range;
    this._id = new Date().getTime();
  }

  /**
   * Gets the name of this named range.
   */
  public getName() : string {
    return this._name;
  }

  /**
   * Gets the range referenced by this named range.
   */
  public getRange() : Range {
    return this._range;
  }

  /**
   * Deletes this named range.
   */
  public remove() : void {
    this._parent.removeNamedRanges(this);
  }

  /**
   * Sets/updates the name of the named range.
   */
  public setName(name: string) : NamedRange {
    this._name = name;
    return this;
  }

  /**
   * Sets/updates the range for this named range.
   */
  public setRange(range: Range) : NamedRange {
    this._range = range;
    return this;
  }

  public get id() : number {
    return this._id;
  }

}
