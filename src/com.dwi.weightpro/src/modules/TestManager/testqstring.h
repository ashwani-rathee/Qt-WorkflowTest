#ifndef TESTQSTRING_H
#define TESTQSTRING_H


#include "TestSuite.h"
#include "qtestcase.h"

class TestQString: public TestSuite
{
    Q_OBJECT

public:
    using TestSuite::TestSuite;

private slots:
    void testFromUtf8();
    void testToUtf8();
};

static TestQString TEST_QSTRING;

#endif // TESTQSTRING_H
