/// <reference types="tree-sitter-cli/dsl" />

// Case-insensitive regex generator
function ci(keyword) {
  const pattern = keyword
    .split("")
    .map((ch) => {
      if (/[A-Za-z]/.test(ch)) {
        return `[${ch.toLowerCase()}${ch.toUpperCase()}]`;
      }
      return ch.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&");
    })
    .join("");
  return new RegExp(pattern);
}

// Case-insensitive keyword token with precedence over identifiers
function kw(word) {
  return alias(token(prec(2, ci(word))), word);
}

function commaSep(rule) {
  return optional(commaSep1(rule));
}

function commaSep1(rule) {
  return seq(rule, repeat(seq(",", rule)));
}

module.exports = grammar({
  name: "vba",

  extras: ($) => [/[ \t\f\u00a0]+/, $._line_continuation],

  externals: ($) => [],

  word: ($) => $.identifier,

  conflicts: ($) => [
    [$._lvalue, $._expression],
    [$.call_statement, $._expression],
    [$.dotted_name, $._expression],
  ],

  rules: {
    source_file: ($) => repeat($._item),

    _item: ($) =>
      choice(
        $._declaration,
        $.comment,
        $._newline,
        $.class_header,
      ),

    _newline: ($) => /\r?\n/,
    _terminator: ($) => choice($._newline, ":"),

    _line_continuation: ($) => token(seq("_", /[ \t]*/, /\r?\n/)),

    // ── Class file header ──────────────────────────────────────────
    // VERSION 1.0 CLASS\nBEGIN\n  ...\nEND\n
    // Captured as a single token since the content is opaque metadata
    class_header: ($) =>
      token(
        seq(
          ci("VERSION"),
          /[ \t]+/,
          /[0-9]+\.[0-9]+/,
          /[ \t]+/,
          ci("CLASS"),
          /\r?\n/,
          ci("BEGIN"),
          /\r?\n/,
          // Match lines until END on its own line
          repeat(seq(/[ \t]+[^\r\n]*/, /\r?\n/)),
          ci("END"),
        ),
      ),

    // ── Top-level declarations ─────────────────────────────────────
    _declaration: ($) =>
      choice(
        $.attribute_statement,
        $.option_statement,
        $.sub_declaration,
        $.function_declaration,
        $.property_declaration,
        $.variable_declaration,
        $.const_declaration,
        $.type_declaration,
        $.enum_declaration,
        $.declare_statement,
        $.preprocessor_if,
        $.preprocessor_const,
      ),

    attribute_statement: ($) =>
      seq(
        kw("Attribute"),
        field("name", $.dotted_name),
        "=",
        field("value", $._expression),
        $._terminator,
      ),

    dotted_name: ($) =>
      prec.left(seq($.identifier, repeat(seq(".", $.identifier)))),

    option_statement: ($) =>
      seq(
        kw("Option"),
        choice(kw("Explicit"), kw("Base"), kw("Compare"), kw("Private")),
        optional(choice($.identifier, $.integer_literal)),
        $._terminator,
      ),

    // ── Procedure declarations ─────────────────────────────────────
    sub_declaration: ($) =>
      seq(
        optional($.access_modifier),
        optional(kw("Static")),
        kw("Sub"),
        field("name", $.identifier),
        optional($.parameter_list),
        $._terminator,
        optional($.body),
        kw("End"),
        kw("Sub"),
        $._terminator,
      ),

    function_declaration: ($) =>
      seq(
        optional($.access_modifier),
        optional(kw("Static")),
        kw("Function"),
        field("name", $.identifier),
        optional($.parameter_list),
        optional($.as_clause),
        $._terminator,
        optional($.body),
        kw("End"),
        kw("Function"),
        $._terminator,
      ),

    property_declaration: ($) =>
      seq(
        optional($.access_modifier),
        optional(kw("Static")),
        kw("Property"),
        field("kind", choice(kw("Get"), kw("Let"), kw("Set"))),
        field("name", $.identifier),
        optional($.parameter_list),
        optional($.as_clause),
        $._terminator,
        optional($.body),
        kw("End"),
        kw("Property"),
        $._terminator,
      ),

    parameter_list: ($) => seq("(", commaSep($.parameter), ")"),

    parameter: ($) =>
      seq(
        optional(choice(kw("Optional"), kw("ParamArray"))),
        optional(choice(kw("ByVal"), kw("ByRef"))),
        field("name", $.identifier),
        optional(seq("(", ")")), // array parameter
        optional($.as_clause),
        optional(seq("=", field("default", $._expression))),
      ),

    as_clause: ($) =>
      seq(kw("As"), optional(kw("New")), field("type", $.type_expression)),

    type_expression: ($) => choice($.builtin_type, $.dotted_name),

    builtin_type: ($) =>
      choice(
        kw("Boolean"),
        kw("Byte"),
        kw("Integer"),
        kw("Long"),
        kw("LongLong"),
        kw("LongPtr"),
        kw("Single"),
        kw("Double"),
        kw("Currency"),
        kw("Decimal"),
        kw("Date"),
        kw("String"),
        kw("Variant"),
        kw("Object"),
      ),

    access_modifier: ($) =>
      choice(kw("Public"), kw("Private"), kw("Friend")),

    // ── Body (sequence of statements) ──────────────────────────────
    body: ($) => repeat1($._statement_with_comment),

    _statement_with_comment: ($) =>
      choice(
        $._statement,
        $.comment,
        $._newline,
      ),

    // ── Statements ─────────────────────────────────────────────────
    _statement: ($) =>
      choice(
        $.variable_declaration,
        $.const_declaration,
        $.set_statement,
        $._call_or_assign,
        $.if_statement,
        $.for_statement,
        $.for_each_statement,
        $.do_loop_statement,
        $.while_wend_statement,
        $.select_case_statement,
        $.with_statement,
        $.on_error_statement,
        $.exit_statement,
        $.goto_statement,
        $.resume_statement,
        $.redim_statement,
        $.erase_statement,
        $.label,
        $.preprocessor_if,
        $.preprocessor_const,
      ),

    variable_declaration: ($) =>
      seq(
        choice(kw("Dim"), kw("Static")),
        commaSep1($.variable_declarator),
        $._terminator,
      ),

    variable_declarator: ($) =>
      seq(
        field("name", $.identifier),
        optional(seq("(", optional($._expression), ")")),
        optional($.as_clause),
      ),

    const_declaration: ($) =>
      seq(
        optional($.access_modifier),
        kw("Const"),
        commaSep1($.const_declarator),
        $._terminator,
      ),

    const_declarator: ($) =>
      seq(
        field("name", $.identifier),
        optional($.as_clause),
        "=",
        field("value", $._expression),
      ),

    // In VBA, assignment and call look the same until you see "=" or not.
    // foo.bar(x) could be a call; foo.bar(x) = y is assignment.
    // We unify them here and let highlights.scm sort it out.
    _call_or_assign: ($) =>
      choice(
        $.assignment_statement,
        $.call_statement,
      ),

    assignment_statement: ($) =>
      seq(
        optional(kw("Let")),
        field("left", $._lvalue),
        "=",
        field("right", $._expression),
        $._terminator,
      ),

    set_statement: ($) =>
      seq(
        kw("Set"),
        field("left", $._lvalue),
        "=",
        field("right", $._expression),
        $._terminator,
      ),

    _lvalue: ($) =>
      choice(
        $.identifier,
        $.member_access_expression,
        $.index_expression,
      ),

    call_statement: ($) =>
      choice(
        // "Call foo(args)" with explicit Call keyword
        seq(
          kw("Call"),
          field("target", $._expression),
          $._terminator,
        ),
        // Bare call with no-parens args: MsgBox "hello", vbInfo
        seq(
          field("target", choice($.identifier, $.member_access_expression)),
          $.argument_list,
          $._terminator,
        ),
        // Bare call with no args (expression as statement): Auto_Open
        seq(
          field("target", choice($.identifier, $.member_access_expression, $.index_expression)),
          $._terminator,
        ),
      ),

    argument_list: ($) =>
      // No-parens arguments: MsgBox "hello", vbInfo
      commaSep1($._argument),

    _argument: ($) =>
      choice(
        // Named argument: param:=value
        seq($.identifier, ":=", $._expression),
        $._expression,
      ),

    // ── Control flow ───────────────────────────────────────────────
    if_statement: ($) =>
      // Block If
      seq(
        kw("If"),
        field("condition", $._expression),
        kw("Then"),
        $._terminator,
        optional($.body),
        repeat($.elseif_clause),
        optional($.else_clause),
        kw("End"),
        kw("If"),
        $._terminator,
      ),

    elseif_clause: ($) =>
      seq(
        kw("ElseIf"),
        field("condition", $._expression),
        kw("Then"),
        $._terminator,
        optional($.body),
      ),

    else_clause: ($) =>
      seq(kw("Else"), $._terminator, optional($.body)),

    for_statement: ($) =>
      seq(
        kw("For"),
        field("variable", $.identifier),
        "=",
        field("start", $._expression),
        kw("To"),
        field("end", $._expression),
        optional(seq(kw("Step"), field("step", $._expression))),
        $._terminator,
        optional($.body),
        kw("Next"),
        optional($.identifier),
        $._terminator,
      ),

    for_each_statement: ($) =>
      seq(
        kw("For"),
        kw("Each"),
        field("variable", $.identifier),
        kw("In"),
        field("collection", $._expression),
        $._terminator,
        optional($.body),
        kw("Next"),
        optional($.identifier),
        $._terminator,
      ),

    do_loop_statement: ($) =>
      choice(
        // Do While/Until ... Loop
        seq(
          kw("Do"),
          choice(kw("While"), kw("Until")),
          field("condition", $._expression),
          $._terminator,
          optional($.body),
          kw("Loop"),
          $._terminator,
        ),
        // Do ... Loop While/Until
        seq(
          kw("Do"),
          $._terminator,
          optional($.body),
          kw("Loop"),
          choice(kw("While"), kw("Until")),
          field("condition", $._expression),
          $._terminator,
        ),
        // Do ... Loop (infinite)
        seq(
          kw("Do"),
          $._terminator,
          optional($.body),
          kw("Loop"),
          $._terminator,
        ),
      ),

    while_wend_statement: ($) =>
      seq(
        kw("While"),
        field("condition", $._expression),
        $._terminator,
        optional($.body),
        kw("Wend"),
        $._terminator,
      ),

    select_case_statement: ($) =>
      seq(
        kw("Select"),
        kw("Case"),
        field("expression", $._expression),
        $._terminator,
        repeat($.case_clause),
        optional($.case_else_clause),
        kw("End"),
        kw("Select"),
        $._terminator,
      ),

    case_clause: ($) =>
      seq(
        kw("Case"),
        commaSep1($.case_expression),
        $._terminator,
        optional($.body),
      ),

    case_expression: ($) =>
      choice(
        seq(kw("Is"), choice("=", "<>", "<", ">", "<=", ">="), $._expression),
        prec.left(seq($._expression, kw("To"), $._expression)),
        $._expression,
      ),

    case_else_clause: ($) =>
      seq(kw("Case"), kw("Else"), $._terminator, optional($.body)),

    with_statement: ($) =>
      seq(
        kw("With"),
        field("object", $._expression),
        $._terminator,
        optional($.body),
        kw("End"),
        kw("With"),
        $._terminator,
      ),

    // ── Error handling ─────────────────────────────────────────────
    on_error_statement: ($) =>
      seq(
        kw("On"),
        kw("Error"),
        choice(
          seq(kw("GoTo"), choice($.identifier, "0", seq("-", "1"))),
          seq(kw("Resume"), kw("Next")),
        ),
        $._terminator,
      ),

    exit_statement: ($) =>
      seq(
        kw("Exit"),
        choice(
          kw("Sub"),
          kw("Function"),
          kw("Property"),
          kw("For"),
          kw("Do"),
        ),
        $._terminator,
      ),

    goto_statement: ($) => seq(kw("GoTo"), $.identifier, $._terminator),

    resume_statement: ($) =>
      seq(kw("Resume"), optional(choice(kw("Next"), $.identifier)), $._terminator),

    label: ($) => prec(1, seq($.identifier, ":")),

    // ── Other statements ───────────────────────────────────────────
    redim_statement: ($) =>
      seq(
        kw("ReDim"),
        optional(kw("Preserve")),
        commaSep1(
          seq(
            field("name", $.identifier),
            "(",
            commaSep1($._expression),
            ")",
            optional($.as_clause),
          ),
        ),
        $._terminator,
      ),

    erase_statement: ($) =>
      seq(kw("Erase"), commaSep1($.identifier), $._terminator),

    // ── Type and Enum declarations ─────────────────────────────────
    type_declaration: ($) =>
      seq(
        optional($.access_modifier),
        kw("Type"),
        field("name", $.identifier),
        $._terminator,
        repeat($.type_member),
        kw("End"),
        kw("Type"),
        $._terminator,
      ),

    type_member: ($) =>
      seq(field("name", $.identifier), $.as_clause, $._terminator),

    enum_declaration: ($) =>
      seq(
        optional($.access_modifier),
        kw("Enum"),
        field("name", $.identifier),
        $._terminator,
        repeat($.enum_member),
        kw("End"),
        kw("Enum"),
        $._terminator,
      ),

    enum_member: ($) =>
      seq(
        field("name", $.identifier),
        optional(seq("=", field("value", $._expression))),
        $._terminator,
      ),

    // ── Declare statement (API declarations) ───────────────────────
    declare_statement: ($) =>
      seq(
        optional($.access_modifier),
        kw("Declare"),
        optional(kw("PtrSafe")),
        choice(kw("Sub"), kw("Function")),
        field("name", $.identifier),
        kw("Lib"),
        field("lib", $.string_literal),
        optional(seq(kw("Alias"), field("alias", $.string_literal))),
        optional($.parameter_list),
        optional($.as_clause),
        $._terminator,
      ),

    // ── Preprocessor directives ────────────────────────────────────
    // Use single tokens for preprocessor keywords to avoid ambiguity
    // with '#' appearing at start of _item
    _pp_if: ($) => token(prec(3, seq("#", /[ \t]*/, ci("If")))),
    _pp_elseif: ($) => token(prec(3, seq("#", /[ \t]*/, ci("ElseIf")))),
    _pp_else: ($) => token(prec(3, seq("#", /[ \t]*/, ci("Else")))),
    _pp_end_if: ($) => token(prec(3, seq("#", /[ \t]*/, ci("End"), /[ \t]+/, ci("If")))),
    _pp_const: ($) => token(prec(3, seq("#", /[ \t]*/, ci("Const")))),

    preprocessor_if: ($) =>
      seq(
        alias($._pp_if, "#If"),
        field("condition", $._expression),
        kw("Then"),
        $._terminator,
        repeat($._item),
        repeat($.preprocessor_elseif),
        optional($.preprocessor_else),
        alias($._pp_end_if, "#End If"),
        $._terminator,
      ),

    preprocessor_elseif: ($) =>
      seq(
        alias($._pp_elseif, "#ElseIf"),
        field("condition", $._expression),
        kw("Then"),
        $._terminator,
        repeat($._item),
      ),

    preprocessor_else: ($) =>
      seq(alias($._pp_else, "#Else"), $._terminator, repeat($._item)),

    preprocessor_const: ($) =>
      seq(
        alias($._pp_const, "#Const"),
        field("name", $.identifier),
        "=",
        field("value", $._expression),
        $._terminator,
      ),

    // ── Expressions ────────────────────────────────────────────────
    _expression: ($) =>
      choice(
        $.literal,
        $.identifier,
        $.member_access_expression,
        $.index_expression,
        $.unary_expression,
        $.binary_expression,
        $.parenthesized_expression,
        $.new_expression,
        $.typeof_expression,
      ),

    parenthesized_expression: ($) => seq("(", $._expression, ")"),

    unary_expression: ($) =>
      prec.right(
        9,
        choice(
          seq("-", $._expression),
          seq(kw("Not"), $._expression),
          seq(kw("AddressOf"), $._expression),
        ),
      ),

    binary_expression: ($) => {
      const table = [
        [8, "^"],
        [7, choice("*", "/")],
        [6, "\\"],
        [5, kw("Mod")],
        [4, choice("+", "-")],
        [3, "&"],
        [
          2,
          choice("=", "<>", "<", ">", "<=", ">=", kw("Is"), kw("Like")),
        ],
        [1, choice(kw("And"), kw("Or"), kw("Xor"), kw("Eqv"), kw("Imp"))],
      ];
      return choice(
        ...table.map(([precedence, operator]) =>
          prec.left(
            precedence,
            seq(
              field("left", $._expression),
              field("operator", operator),
              field("right", $._expression),
            ),
          ),
        ),
      );
    },

    member_access_expression: ($) =>
      prec.left(
        11,
        seq(
          optional(field("object", $._expression)),
          ".",
          field("member", $.identifier),
        ),
      ),

    index_expression: ($) =>
      prec.left(
        11,
        seq(
          field("object", $._expression),
          "(",
          commaSep($._expression),
          ")",
        ),
      ),

    new_expression: ($) => seq(kw("New"), $.type_expression),

    typeof_expression: ($) =>
      prec.right(
        2,
        seq(kw("TypeOf"), $._expression, kw("Is"), $.type_expression),
      ),

    // ── Literals ───────────────────────────────────────────────────
    literal: ($) =>
      choice(
        $.integer_literal,
        $.float_literal,
        $.string_literal,
        $.boolean_literal,
        $.nothing_literal,
        $.date_literal,
      ),

    integer_literal: ($) =>
      token(
        choice(
          /[0-9]+[&%^]?/,
          /&[hH][0-9a-fA-F]+&?/,
          /&[oO][0-7]+&?/,
        ),
      ),

    float_literal: ($) =>
      token(/[0-9]*\.[0-9]+([eE][+-]?[0-9]+)?[!#@]?/),

    string_literal: ($) =>
      token(seq('"', repeat(choice(/[^"\r\n]/, '""')), '"')),

    boolean_literal: ($) => choice(kw("True"), kw("False")),

    nothing_literal: ($) => kw("Nothing"),

    date_literal: ($) => token(seq("#", /[^#\r\n]+/, "#")),

    // ── Terminals ──────────────────────────────────────────────────
    comment: ($) =>
      token(
        choice(
          seq("'", /[^\r\n]*/),
          seq(ci("Rem"), /[ \t]/, /[^\r\n]*/),
        ),
      ),

    identifier: ($) => /[A-Za-z_][A-Za-z0-9_]*[%&!#$@]?/,
  },
});
