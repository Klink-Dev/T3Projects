1. Pip install the following packages:

  a. urllib3

  b. requests

  c. requests_kerberos

  d. pendulum

  e. pandas

  f. tabulate

  g. apscheduler

(Click the Terminal tab at the bottom of pycharm and type pip install <package name>)

2. Change the Fc name in line 268 from RNO4 to your FC

3. Line 140(ish) or self.process_ids is the ppr lines that it looks at for zero processed units (Add or remove for specific departments)

4. Line 149(ish) or self.exclude_functions is the names of the purely indirect process that you want to exclude. Note they have to be exact as the CSV shows.

5. Line 152(ish) or self.url is the url of the webhook that you are sending the alerts to

6. Run the file
