export declare type TokenType = "FUNCTION" | "STRING" | "NUMBER" | "UNARY_MINUS" | "BOOLEAN" | "WORD" | "OPERATOR" | "ARRAY" | "VARIABLE";
export declare class Token {
    type: TokenType;
    value: string | number;
    constructor(tokenType: TokenType, value: string | number);
    toString(): string;
}
export declare type ArrayNodeValue = (string | number | ArrayNodeValue)[];
export declare class ArrayLiteralNode {
    elements: ArrayNodeValue;
    constructor(elements: ArrayNodeValue);
    evaluate(): ArrayNodeValue;
}
export declare type ASTNode = {
    type: string;
    value?: string | number;
    operands?: (ASTNode | ArrayLiteralNode)[];
};
export declare type Context = {
    [key: string]: boolean | number | object | string | undefined;
};
export declare const ValidFuncNames: string[];
//# sourceMappingURL=FormulaEvaluation.types.d.ts.map