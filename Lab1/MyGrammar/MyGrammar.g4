grammar MyGrammar;

options {
  language = CSharp;
}

expression 
  :  left=expression ('*' | '/' | '^' | 'div' | 'mod') right=expression      #MultiplicativeOrPowerExpression
  |  left=expression ('+' | '-') right=expression            #AdditiveExpression
  |  'max' '(' exprList ')'                                  #MaxFunction
  |  'min' '(' exprList ')'                                  #MinFunction
  |  '(' expression ')'                                      #ParenthesizedExpression
  |  NUMBER                                                  #Number
  |  CELL                                                    #Cell
  ;

exprList: expression (',' expression)*;

CELL: [A-Z]+ [1-9][0-9]*;


NUMBER: ('-')?[0-9]+('.'[0-9]+)?;

WS: [ \t\r\n] -> skip;

INVALID: .;
