# ProvisioningRequests
Tools for provisioning requests.

A Microsoft Flow reads an excel sheet of all the provisioning requests every 2 weeks, and sends them to the provsioningPR.cs Azure Function. This function updates the audience config json with each request, noting successes and failures. It then creates a PR for the change, and a workitem with all the things that were done.

Another webhook activates when the workitem is closed (when the PR is completed). This reads a series of tables in the workitem and notifies each request submitter of the results of their requests.
