/* Generated file, do not edit */

#ifndef CXXTEST_RUNNING
#define CXXTEST_RUNNING
#endif

#define _CXXTEST_HAVE_STD
#define _CXXTEST_HAVE_EH
#include <cxxtest/TestListener.h>
#include <cxxtest/TestTracker.h>
#include <cxxtest/TestRunner.h>
#include <cxxtest/RealDescriptions.h>
#include <cxxtest/TestMain.h>
#include <cxxtest/ParenPrinter.h>

int main( int argc, char *argv[] ) {
 int status;
    CxxTest::ParenPrinter tmp;
    CxxTest::RealWorldDescription::_worldName = "cxxtest";
    status = CxxTest::Main< CxxTest::ParenPrinter >( tmp, argc, argv );
    return status;
}
bool suite_IntegrationTestSuite_init = false;
#include "C:\Users\taf656\Development\xlnt\source\tests\integration\IntegrationTestSuite.h"

static IntegrationTestSuite suite_IntegrationTestSuite;

static CxxTest::List Tests_IntegrationTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_IntegrationTestSuite( "../../source/tests/integration/IntegrationTestSuite.h", 9, "IntegrationTestSuite", suite_IntegrationTestSuite, Tests_IntegrationTestSuite );

static class TestDescription_suite_IntegrationTestSuite_test_1 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IntegrationTestSuite_test_1() : CxxTest::RealTestDescription( Tests_IntegrationTestSuite, suiteDescription_IntegrationTestSuite, 17, "test_1" ) {}
 void runTest() { suite_IntegrationTestSuite.test_1(); }
} testDescription_suite_IntegrationTestSuite_test_1;

#include "C:\Users\taf656\Development\xlnt\source\tests\packaging\NullableTestSuite.h"

static NullableTestSuite suite_NullableTestSuite;

static CxxTest::List Tests_NullableTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_NullableTestSuite( "../../source/tests/packaging/NullableTestSuite.h", 8, "NullableTestSuite", suite_NullableTestSuite, Tests_NullableTestSuite );

static class TestDescription_suite_NullableTestSuite_test_has_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_has_value() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 16, "test_has_value" ) {}
 void runTest() { suite_NullableTestSuite.test_has_value(); }
} testDescription_suite_NullableTestSuite_test_has_value;

static class TestDescription_suite_NullableTestSuite_test_get_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_get_value() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 22, "test_get_value" ) {}
 void runTest() { suite_NullableTestSuite.test_get_value(); }
} testDescription_suite_NullableTestSuite_test_get_value;

static class TestDescription_suite_NullableTestSuite_test_equality : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_equality() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 29, "test_equality" ) {}
 void runTest() { suite_NullableTestSuite.test_equality(); }
} testDescription_suite_NullableTestSuite_test_equality;

static class TestDescription_suite_NullableTestSuite_test_assignment : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_assignment() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 36, "test_assignment" ) {}
 void runTest() { suite_NullableTestSuite.test_assignment(); }
} testDescription_suite_NullableTestSuite_test_assignment;

static class TestDescription_suite_NullableTestSuite_test_copy_constructor : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_copy_constructor() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 44, "test_copy_constructor" ) {}
 void runTest() { suite_NullableTestSuite.test_copy_constructor(); }
} testDescription_suite_NullableTestSuite_test_copy_constructor;

#include "C:\Users\taf656\Development\xlnt\source\tests\packaging\PackageTestSuite.h"

static ZipPackageTestSuite suite_ZipPackageTestSuite;

static CxxTest::List Tests_ZipPackageTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_ZipPackageTestSuite( "../../source/tests/packaging/PackageTestSuite.h", 9, "ZipPackageTestSuite", suite_ZipPackageTestSuite, Tests_ZipPackageTestSuite );

static class TestDescription_suite_ZipPackageTestSuite_test_read_text : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipPackageTestSuite_test_read_text() : CxxTest::RealTestDescription( Tests_ZipPackageTestSuite, suiteDescription_ZipPackageTestSuite, 17, "test_read_text" ) {}
 void runTest() { suite_ZipPackageTestSuite.test_read_text(); }
} testDescription_suite_ZipPackageTestSuite_test_read_text;

static class TestDescription_suite_ZipPackageTestSuite_test_write_text : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipPackageTestSuite_test_write_text() : CxxTest::RealTestDescription( Tests_ZipPackageTestSuite, suiteDescription_ZipPackageTestSuite, 29, "test_write_text" ) {}
 void runTest() { suite_ZipPackageTestSuite.test_write_text(); }
} testDescription_suite_ZipPackageTestSuite_test_write_text;

static class TestDescription_suite_ZipPackageTestSuite_test_read_xml : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipPackageTestSuite_test_read_xml() : CxxTest::RealTestDescription( Tests_ZipPackageTestSuite, suiteDescription_ZipPackageTestSuite, 54, "test_read_xml" ) {}
 void runTest() { suite_ZipPackageTestSuite.test_read_xml(); }
} testDescription_suite_ZipPackageTestSuite_test_read_xml;

#include "C:\Users\taf656\Development\xlnt\source\tests\packaging\UriTestSuite.h"

static uriTestSuite suite_uriTestSuite;

static CxxTest::List Tests_uriTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_uriTestSuite( "../../source/tests/packaging/UriTestSuite.h", 6, "uriTestSuite", suite_uriTestSuite, Tests_uriTestSuite );

static class TestDescription_suite_uriTestSuite_test_absolute_path : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_absolute_path() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 14, "test_absolute_path" ) {}
 void runTest() { suite_uriTestSuite.test_absolute_path(); }
} testDescription_suite_uriTestSuite_test_absolute_path;

static class TestDescription_suite_uriTestSuite_test_absolute_uri : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_absolute_uri() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 19, "test_absolute_uri" ) {}
 void runTest() { suite_uriTestSuite.test_absolute_uri(); }
} testDescription_suite_uriTestSuite_test_absolute_uri;

static class TestDescription_suite_uriTestSuite_test_authority : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_authority() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 24, "test_authority" ) {}
 void runTest() { suite_uriTestSuite.test_authority(); }
} testDescription_suite_uriTestSuite_test_authority;

static class TestDescription_suite_uriTestSuite_test_dns_safe_host : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_dns_safe_host() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 29, "test_dns_safe_host" ) {}
 void runTest() { suite_uriTestSuite.test_dns_safe_host(); }
} testDescription_suite_uriTestSuite_test_dns_safe_host;

static class TestDescription_suite_uriTestSuite_test_fragment : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_fragment() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 34, "test_fragment" ) {}
 void runTest() { suite_uriTestSuite.test_fragment(); }
} testDescription_suite_uriTestSuite_test_fragment;

static class TestDescription_suite_uriTestSuite_test_get_host : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_get_host() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 39, "test_get_host" ) {}
 void runTest() { suite_uriTestSuite.test_get_host(); }
} testDescription_suite_uriTestSuite_test_get_host;

static class TestDescription_suite_uriTestSuite_test_host_name_type : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_host_name_type() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 44, "test_host_name_type" ) {}
 void runTest() { suite_uriTestSuite.test_host_name_type(); }
} testDescription_suite_uriTestSuite_test_host_name_type;

static class TestDescription_suite_uriTestSuite_test_is_absolute_uri : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_is_absolute_uri() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 49, "test_is_absolute_uri" ) {}
 void runTest() { suite_uriTestSuite.test_is_absolute_uri(); }
} testDescription_suite_uriTestSuite_test_is_absolute_uri;

static class TestDescription_suite_uriTestSuite_test_default_port : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_default_port() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 54, "test_default_port" ) {}
 void runTest() { suite_uriTestSuite.test_default_port(); }
} testDescription_suite_uriTestSuite_test_default_port;

static class TestDescription_suite_uriTestSuite_test_is_file : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_is_file() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 59, "test_is_file" ) {}
 void runTest() { suite_uriTestSuite.test_is_file(); }
} testDescription_suite_uriTestSuite_test_is_file;

static class TestDescription_suite_uriTestSuite_test_is_loopback : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_is_loopback() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 64, "test_is_loopback" ) {}
 void runTest() { suite_uriTestSuite.test_is_loopback(); }
} testDescription_suite_uriTestSuite_test_is_loopback;

static class TestDescription_suite_uriTestSuite_test_is_unc : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_is_unc() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 69, "test_is_unc" ) {}
 void runTest() { suite_uriTestSuite.test_is_unc(); }
} testDescription_suite_uriTestSuite_test_is_unc;

static class TestDescription_suite_uriTestSuite_test_local_path : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_local_path() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 74, "test_local_path" ) {}
 void runTest() { suite_uriTestSuite.test_local_path(); }
} testDescription_suite_uriTestSuite_test_local_path;

static class TestDescription_suite_uriTestSuite_test_original_string : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_original_string() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 79, "test_original_string" ) {}
 void runTest() { suite_uriTestSuite.test_original_string(); }
} testDescription_suite_uriTestSuite_test_original_string;

static class TestDescription_suite_uriTestSuite_test_path_and_query : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_path_and_query() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 84, "test_path_and_query" ) {}
 void runTest() { suite_uriTestSuite.test_path_and_query(); }
} testDescription_suite_uriTestSuite_test_path_and_query;

static class TestDescription_suite_uriTestSuite_test_port : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_port() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 89, "test_port" ) {}
 void runTest() { suite_uriTestSuite.test_port(); }
} testDescription_suite_uriTestSuite_test_port;

static class TestDescription_suite_uriTestSuite_test_query : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_query() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 94, "test_query" ) {}
 void runTest() { suite_uriTestSuite.test_query(); }
} testDescription_suite_uriTestSuite_test_query;

static class TestDescription_suite_uriTestSuite_test_scheme : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_scheme() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 99, "test_scheme" ) {}
 void runTest() { suite_uriTestSuite.test_scheme(); }
} testDescription_suite_uriTestSuite_test_scheme;

static class TestDescription_suite_uriTestSuite_test_user_escaped : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_user_escaped() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 104, "test_user_escaped" ) {}
 void runTest() { suite_uriTestSuite.test_user_escaped(); }
} testDescription_suite_uriTestSuite_test_user_escaped;

static class TestDescription_suite_uriTestSuite_test_user_info : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_uriTestSuite_test_user_info() : CxxTest::RealTestDescription( Tests_uriTestSuite, suiteDescription_uriTestSuite, 109, "test_user_info" ) {}
 void runTest() { suite_uriTestSuite.test_user_info(); }
} testDescription_suite_uriTestSuite_test_user_info;

#include <cxxtest/Root.cpp>
const char* CxxTest::RealWorldDescription::_worldName = "cxxtest";
