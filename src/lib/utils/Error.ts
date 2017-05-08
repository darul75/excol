export class Errors {
  public static NOT_IMPLEMENTED_YET = 'Not implemented yet';
  public static INCORRECT_GET_RANGE_USAGE = 'Please use getRanges() method instead.';
  public static INCORRECT_RANGE_SIGNATURE = 'Please provide row, column at least.';
  public static INCORRECT_A1_NOTATION = 'Please provide correct A1 notation.';
  public static INCORRECT_MERGE = 'You must select all cells in a merged range to merge or unmerge them.';
  public static INCORRECT_MERGE_SINGLE_CELL = 'Can not merge single cell.';
  public static INCORRECT_MERGE_HORIZONTAL = "You can't merge horizontally across an existing vertically merged section.";
  public static INCORRECT_MERGE_VERTICALLY = "You can't merge vertically across an existing horizontally merged section.";
  public static INCORRECT_COORDINATES_OR_DIMENSIONS = 'The coordinates or dimensions of the range are invalid.';
  public static INCORRECT_MOVE_TO = `You can't perform a cut/paste from a range that partially intersects a merge. Consider unmerging or selecting a larger range that includes the entire merge.`;
  public static INCORRECT_RANGE_HEIGHT = (first, second) => `Incorrect range height, was ${first} but should be ${second}.`;
  public static INCORRECT_RANGE_WIDTH = (first, second) => `Incorrect range width, was ${first} but should be ${second}.`;
  public static INCORRECT_SET_VALUES_PARAMETERS = (type) => `Cannot find method setValues(${type}).`;
  public static INCORRECT_RANGE_DATA_VALIDATION = (a1Notation) => `The data you entered in cell(s) ${a1Notation} violates the data validation rules set on this cell.`;
}
