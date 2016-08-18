// Interfaces
export interface A1Type {
  A1: string;
}

export interface MatrixConverterFunc {
  (A1: A1Type): string[][];
}

export interface CoordinatesConverterFunc {
  (A1: A1Type): any[];
}