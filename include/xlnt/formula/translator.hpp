#include <string>
#include <vector>

namespace xlnt {

class cell_reference;
class tokenizer;

class translator
{
    translator(const std::string &formula, const cell_reference &ref);

    std::vector<std::string> get_tokens();

    static std::string translate_row(const std::string &row_str, int row_delta);
    static std::string translate_col(const std::string &col_str, col_delta);

    std::pair<std::string, std::string> strip_ws_name(const std::string &range_str);

    void translate_range(const range_reference &range_ref);
    void translate_formula(const cell_reference &dest);

  private:
    const std::string ROW_RANGE_RE;
    const std::string COL_RANGE_RE;
    const std::string CELL_REF_RE;

    std::string formula_;
    cell_reference reference_;
    tokenizer tokenizer_;
};

} // namespace xlnt
