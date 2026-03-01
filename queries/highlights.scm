; Comments
(comment) @comment

; Literals
(string_literal) @string
(integer_literal) @number
(float_literal) @number
(boolean_literal) @boolean
(nothing_literal) @constant.builtin
(date_literal) @string.special

; Built-in types
(builtin_type) @type.builtin

; Procedure declarations
(sub_declaration
  name: (identifier) @function)
(function_declaration
  name: (identifier) @function)
(property_declaration
  name: (identifier) @function)
(declare_statement
  name: (identifier) @function)

; Parameters
(parameter
  name: (identifier) @variable.parameter)

; Variable/const declarations
(variable_declarator
  name: (identifier) @variable)
(const_declarator
  name: (identifier) @constant)

; Type and Enum declarations
(type_declaration
  name: (identifier) @type)
(enum_declaration
  name: (identifier) @type)
(enum_member
  name: (identifier) @constant)
(type_member
  name: (identifier) @variable.member)

; Member access
(member_access_expression
  member: (identifier) @variable.member)

; Function calls
(call_statement
  target: (identifier) @function.call)
(call_statement
  target: (member_access_expression
    member: (identifier) @function.call))

; Index expressions (array access / function call with parens)
(index_expression
  object: (identifier) @function.call)

; Labels
(label
  (identifier) @label)

; Attribute statements
(attribute_statement) @attribute

; Keywords - procedures
[
  "Sub"
  "Function"
  "Property"
] @keyword.function

; Keywords - end/exit
[
  "End"
  "Exit"
] @keyword

; Keywords - declarations
[
  "Dim"
  "Static"
  "Const"
  "Type"
  "Enum"
  "Declare"
  "Lib"
  "Alias"
  "PtrSafe"
  "Option"
  "Attribute"
  "ReDim"
  "Preserve"
  "Erase"
  "Explicit"
  "Base"
  "Compare"
] @keyword

; Keywords - control flow
[
  "If"
  "Then"
  "Else"
  "ElseIf"
  "Select"
  "Case"
] @keyword.conditional

; Keywords - loops
[
  "For"
  "To"
  "Step"
  "Next"
  "Each"
  "In"
  "Do"
  "Loop"
  "While"
  "Until"
  "Wend"
] @keyword.repeat

; Keywords - error handling
[
  "On"
  "Error"
  "Resume"
  "GoTo"
] @keyword.exception

; Keywords - visibility
[
  "Public"
  "Private"
  "Friend"
] @keyword

; Keywords - parameter/type
[
  "ByVal"
  "ByRef"
  "Optional"
  "ParamArray"
  "As"
  "New"
  "Set"
  "Let"
  "Call"
  "With"
  "Get"
] @keyword

; Keywords - operators
[
  "And"
  "Or"
  "Not"
  "Xor"
  "Mod"
  "Is"
  "Like"
  "Eqv"
  "Imp"
  "AddressOf"
  "TypeOf"
] @keyword.operator

; Keywords - boolean
[
  "True"
  "False"
] @boolean

"Nothing" @constant.builtin

; Preprocessor
"#If" @keyword.directive
"#ElseIf" @keyword.directive
"#Else" @keyword.directive
"#End If" @keyword.directive
"#Const" @keyword.directive

; Operators
[
  "="
  "<>"
  "<"
  ">"
  "<="
  ">="
  "+"
  "-"
  "*"
  "/"
  "\\"
  "^"
  "&"
  ":="
] @operator

; Punctuation
[
  "("
  ")"
] @punctuation.bracket

"," @punctuation.delimiter
":" @punctuation.delimiter
"." @punctuation.delimiter
