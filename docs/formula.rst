Parsing Formulas
================

`xlnt` supports limited parsing of formulas embedded in cells. The
`xlnt/formula` module contains a `tokenizer` class to break
formulas into their consitutuent tokens. Usage is as follows:

.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        std::string formula = R"(=IF($A$1,"then True",MAX(DEFAULT_VAL,'Sheet 2'!B1)))";
	xlnt::tokenizer tok(formula);
	tok.parse();

	for(auto &t : tok.get_items())
	{
            std::cout << t.get_value() << "\t" << t.get_type() << "\t" << t.get_subtype() << std::endl;
	}

	// prints:
	//
	// IF(     FUNC    OPEN
	// $A$1    OPERAND RANGE
	// ,       SEP     ARG
	// "then True"     OPERAND TEXT
	// ,       SEP     ARG
        // MAX(    FUNC    OPEN
	// DEFAULT_VAL     OPERAND RANGE
	// ,       SEP     ARG
	// 'Sheet 2'!B1    OPERAND RANGE
	// )       FUNC    CLOSE
	// )       FUNC    CLOSE

        return 0;
    }

As shown above, tokens have three attributes of interest:

* ``.value``: The substring of the formula that produced this token

* ``.type``: The type of token this represents. Can be one of

  - ``token::literal``: If the cell does not contain a formula, its
    value is represented by a single ``LITERAL`` token.

  - ``token::operand``: A generic term for any value in the Excel
    formula. (See ``.subtype`` below for more details).

  - ``token::func``: Function calls are broken up into tokens for the
    opener (e.g., ``SUM(``), followed by the arguments, followed by
    the closer (i.e., ``)``). The function name and opening
    parenthesis together form one ``FUNC`` token, and the matching
    parenthesis forms another ``FUNC`` token.

  - ``token::array``: Array literals (enclosed between curly braces)
    get two ``ARRAY`` tokens each, one for the opening ``{`` and one
    for the closing ``}``.

  - ``token::paren``: When used for grouping subexpressions (and not to
    denote function calls), parentheses are tokenized as ``paren``
    tokens (one per character).

  - ``token::sep``: These tokens are created from either commas (``,``)
    or semicolons (``;``). Commas create ``sep`` tokens when they are
    used to separate function arguments (e.g., ``SUM(a,b)``) or when
    they are used to separate array elements (e.g., ``{a,b}``). (They
    have another use as an infix operator for joining
    ranges). Semicolons are always used to separate rows in an array
    literal, so always create ``sep`` tokens.

  - ``token::op_pre``: Designates a prefix unary operator. Its value is
    always ``+`` or ``-``

  - ``token::op_in``: Designates an infix binary operator. Possible
    values are ``>=``, ``<=``, ``<>``, ``=``, ``>``, ``<``, ``*``,
    ``/``, ``+``, ``-``, ``^``, or ``&``.

  - ``token::op_post``: Designates a postfix unary operator. Its value
    is always ``%``.

  - ``token::wspace``: Created for any whitespace encountered. Its
    value is always a single space, regardless of how much whitespace
    is found.

* ``.subtype``: Some of the token types above use the subtype to
  provide additional information about the token. Possible subtypes
  are:

  + ``token::text``, ``token::number``, ``token::logical``,
    ``token::error``, ``token::range``: these subtypes describe the
    various forms of ``operand`` found in formulae. ``logical`` is
    either ``true`` or ``false``, ``range`` is either a named range or
    a direct reference to another range. ``text``, ``number``, and
    ``error`` all refer to literal values in the formula

  + ``token::open`` and ``token::close``: these two subtypes are used by
    ``paren``, ``func``, and ``array``, to describe whether the token
    is opening a new subexpression or closing it.

  + ``token::arg`` and ``token::row``: are used by the ``sep`` tokens,
    to distinguish between the comma and semicolon. Commas produce
    tokens of subtype ``arg`` whereas semicolons produce tokens of
    subtype ``row``
