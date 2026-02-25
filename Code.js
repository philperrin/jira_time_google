<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
    body {
      font-family: sans-serif;
    }
  </style>
  </head>
  <body>
    <p>Use this form to create a new Jira issue - it will be assigned to you in the project specified.</p>
    <form id="newJira">
      <label for="input1">Project:</label><br>
        <select id="input1" name="input1">

        </select><br><br>
      <label for="input2">Technology:</label><br>
        <select id="input2" name="input2">
          <option value=""></option>
          <option value="Core">Core</option>
          <option value="AIMS">AIMS</option>
          <option value="Atlan">Atlan</option>
          <option value="AWS Glue">AWS Glue</option>
          <option value="dbt">dbt</option>
          <option value="DQLabs">DQLabs</option>
          <option value="HeyWM">HeyWM</option>
          <option value="HVR">HVR</option>
          <option value="InfoSphere">InfoSphere</option>
          <option value="Matillion">Matillion</option>
          <option value="Newton Insights">Newton Insights</option>
          <option value="Power Apps">Power Apps</option>
          <option value="Power Automate">Power Automate</option>
          <option value="Power BI">Power BI</option>
          <option value="Qlik">Qlik</option>
          <option value="Sigma">Sigma</option>
          <option value="Snowflake">Snowflake</option>
          <option value="Spotfire">Spotfire</option>
          <option value="Tableau">Tableau</option>
        </select><br><br>
      <label for="input3">EOps Type:</label><br>
        <select id="input3" name="input3">
          <option value="Core">Core</option>
          <option value="AdvOps">AdvOps</option>
        </select><br><br>
      <label for="input4">Summary:</label><br>
      <input type="text" id="input4" name="input4"><br><br>
      <label for="input5">Description:</label><br>
      <textarea id="input5" name="notes" rows="4" cols="50"></textarea><br><br>
      <input type="button" id="submitBtn" value="Submit" onclick="submitForm(this)">
    </form>
    
    <script>
      // Populate the Project dropdown from the Allocation sheet on load
      google.script.run.withSuccessHandler(populateDropdown).getDropdownValues();

      function populateDropdown(items) {
        var select = document.getElementById('input1');
        select.innerHTML = '';
        items.forEach(function(item) {
          var option = document.createElement('option');
          option.value = item;
          option.text = item;
          select.appendChild(option);
        });
      }

      function submitForm(btn) {
        btn.disabled = true;
        btn.value = 'Submitting…';
        google.script.run
          .withSuccessHandler(function(msg) {
            alert(msg);
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            alert('Error: ' + err.message);
            btn.disabled = false;
            btn.value = 'Submit';
          })
          .makeJira(document.getElementById('newJira'));
      }
    </script>
  </body>
</html>
