#pragma once

#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class uriTestSuite : public CxxTest::TestSuite
{
public:
    uriTestSuite() : uri(complex_url)
    {

    }

    void test_absolute_path()
    {
        TS_ASSERT_EQUALS(uri.get_AbsolutePath(), "/tfussell/EPPlusPlus.git");
    }

    void test_absolute_uri()
    {
        TS_ASSERT_EQUALS(uri.get_Absoluteuri(), complex_url);
    }

    void test_authority()
    {
        TS_ASSERT_EQUALS(uri.get_Authority(), "thomas:password@github.com:71");
    }

    void test_dns_safe_host()
    {
        TS_ASSERT_EQUALS(uri.get_DnsSafeHost(), "github.com");
    }

    void test_fragment()
    {
        TS_ASSERT_EQUALS(uri.get_Fragment(), "abc");
    }

    void test_get_host()
    {
        TS_ASSERT_EQUALS(uri.get_Host(), "github.com");
    }

    void test_host_name_type()
    {
        TS_ASSERT_EQUALS(uri.get_HostNameType(), xlnt::uri_host_name_type::Dns);
    }

    void test_is_absolute_uri()
    {
        TS_ASSERT_EQUALS(uri.IsAbsoluteuri(), true);
    }

    void test_default_port()
    {
        TS_ASSERT_EQUALS(uri.IsDefaultPort(), false);
    }

    void test_is_file()
    {
        TS_ASSERT_EQUALS(uri.IsFile(), false);
    }

    void test_is_loopback()
    {
        TS_ASSERT_EQUALS(uri.IsLoopback(), false);
    }

    void test_is_unc()
    {
        TS_ASSERT_EQUALS(uri.IsUnc(), false);
    }

    void test_local_path()
    {
        TS_ASSERT_EQUALS(uri.get_LocalPath(), "/tfussell/EPPlusPlus.git");
    }

    void test_original_string()
    {
        TS_ASSERT_EQUALS(uri.get_OriginalString(), complex_url);
    }

    void test_path_and_query()
    {
        TS_ASSERT_EQUALS(uri.get_PathAndQuery(), "/tfussell/EPPlusPlus.git?a=1&b=2");
    }

    void test_port()
    {
        TS_ASSERT_EQUALS(uri.get_Port(), 71);
    }

    void test_query()
    {
        TS_ASSERT_EQUALS(uri.get_Query(), "a=1&b=2");
    }

    void test_scheme()
    {
        TS_ASSERT_EQUALS(uri.get_Scheme(), "https");
    }

    void test_user_escaped()
    {
        TS_ASSERT_EQUALS(uri.get_UserEscaped(), false);
    }

    void test_user_info()
    {
        TS_ASSERT_EQUALS(uri.get_UserInfo(), "thomas:password");
    }

private:
    const std::string complex_url = "https://thomas:password@github.com:71/tfussell/EPPlusPlus.git?a=1&b=2#abc";
    xlnt::uri uri;
};
