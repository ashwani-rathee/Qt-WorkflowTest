#ifndef TESTSUITE_H
#define TESTSUITE_H

#include <QObject>
#include <vector>

/** \brief Base class for the test suite runner.
 */
class TestSuite: public QObject
{
public:
     TestSuite();

     static std::vector<QObject*> & suite();
};

#endif // TESTSUITE_H
