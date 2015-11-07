#pragma once

#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class cell_reference;
class tokenizer;

class XLNT_CLASS translator
{
    translator(const string &formula, const cell_reference &ref);

    std::vector<string> get_tokens();

    static string translate_row(const string &row_str, int row_delta);
    static string translate_col(const string &col_str, col_delta);

    std::pair<string, string> strip_ws_name(const string &range_str);

    void translate_range(const range_reference &range_ref);
    void translate_formula(const cell_reference &dest);

  private:
    const string ROW_RANGE_RE;
    const string COL_RANGE_RE;
    const string CELL_REF_RE;

    string formula_;
    cell_reference reference_;
    tokenizer tokenizer_;
};

} // namespace xlnt
