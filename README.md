# Qt-WorkflowTest

Shows how to set up an online installer.

The example uses a very simple web server shipped with python.

Generate online repository with

  repogen -p packages repository

Generate installer with

  binarycreator --online-only -c config/config.xml -p packages installer

Now launch a minimal web server in the example's directory (admin rights may be needed)

  python -m SimpleHTTPServer 80

This should make the content of the local directory available under
http://localhost

You should be able to now launch the installer.

To deploy an update, run

  repogen --update-new-components -p packages_update repository

and launch the maintenance tool in your installation.
