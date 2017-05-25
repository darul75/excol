// Imports

// Class & methods

const CR_STRING = "\n";

/**
 * Logger class.
 */
export class Log {
  private _logs: string[] = [];

  /**
   * Clears the log.
   */
  public clear() {
    this._logs = [];
  }

  /**
   * Returns a complete list of messages in the current log.
   */
  public getLog() : string {
    return this._logs.join(CR_STRING);
  }

  /**
   * Writes the string to the logging console.
   *
   * or
   *
   * Writes a formatted string to the logging console, using the format and values provided.
   *
   * TODO: finish it..
   *
   * @param format
   * @param values
   * @returns {Logger}
   */
  public log(format: string, ...values: any[]) : Log {
    if (values) {
      this.logMulti(format, values);
    }
    else {
      this.logSingle(format);
    }

    console.log(this.getLog());

    return this;
  }

  private logSingle(format: string) : void {
    this.appendLog(format);
  }

  private logMulti(format: string, ...values: any[]) : void {
    this.appendLog(formatForParams(format)(values));
  }

  private appendLog(log: string) {
    this._logs.push(log);
  }

}

// First, checks if it isn't implemented yet.
const formatForParams = (format: string) => {
  return function(o?) {
    const args = arguments;

    const values = args['0'][0];
    let i = 0;
    return format.replace(/%s/g, function(match, number) {
      const val = values[i];

      i++;
      return typeof val != 'undefined'
        ? val
        : match
        ;
    });
  };
};

export const Logger = new Log();
