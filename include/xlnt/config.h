#pragma once

namespace xlnt {

enum class limit_style
{
    openpyxl,
    excel,
    maximum
};

const limit_style LimitStyle = limit_style::openpyxl;

}
