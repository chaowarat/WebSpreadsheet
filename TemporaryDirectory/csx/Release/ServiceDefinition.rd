<?xml version="1.0" encoding="utf-8"?>
<serviceModel xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" name="TemporaryDirectory" generation="1" functional="0" release="0" Id="3ce90da5-19dc-43b3-bfb2-54656ecb94b0" dslVersion="1.2.0.0" xmlns="http://schemas.microsoft.com/dsltools/RDSM">
  <groups>
    <group name="TemporaryDirectoryGroup" generation="1" functional="0" release="0">
      <componentports>
        <inPort name="Spreadsheet:Endpoint1" protocol="http">
          <inToChannel>
            <lBChannelMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/LB:Spreadsheet:Endpoint1" />
          </inToChannel>
        </inPort>
      </componentports>
      <settings>
        <aCS name="Spreadsheet:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" defaultValue="">
          <maps>
            <mapMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/MapSpreadsheet:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" />
          </maps>
        </aCS>
        <aCS name="Spreadsheet:TmpDir" defaultValue="">
          <maps>
            <mapMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/MapSpreadsheet:TmpDir" />
          </maps>
        </aCS>
        <aCS name="SpreadsheetInstances" defaultValue="[1,1,1]">
          <maps>
            <mapMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/MapSpreadsheetInstances" />
          </maps>
        </aCS>
      </settings>
      <channels>
        <lBChannel name="LB:Spreadsheet:Endpoint1">
          <toPorts>
            <inPortMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/Spreadsheet/Endpoint1" />
          </toPorts>
        </lBChannel>
      </channels>
      <maps>
        <map name="MapSpreadsheet:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" kind="Identity">
          <setting>
            <aCSMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/Spreadsheet/Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" />
          </setting>
        </map>
        <map name="MapSpreadsheet:TmpDir" kind="Identity">
          <setting>
            <aCSMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/Spreadsheet/TmpDir" />
          </setting>
        </map>
        <map name="MapSpreadsheetInstances" kind="Identity">
          <setting>
            <sCSPolicyIDMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/SpreadsheetInstances" />
          </setting>
        </map>
      </maps>
      <components>
        <groupHascomponents>
          <role name="Spreadsheet" generation="1" functional="0" release="0" software="C:\Users\CCC-PC\Documents\GitHub\WebSpreadsheet\TemporaryDirectory\csx\Release\roles\Spreadsheet" entryPoint="base\x64\WaHostBootstrapper.exe" parameters="base\x64\WaIISHost.exe " memIndex="1792" hostingEnvironment="frontendadmin" hostingEnvironmentVersion="2">
            <componentports>
              <inPort name="Endpoint1" protocol="http" portRanges="80" />
            </componentports>
            <settings>
              <aCS name="Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" defaultValue="" />
              <aCS name="TmpDir" defaultValue="" />
              <aCS name="__ModelData" defaultValue="&lt;m role=&quot;Spreadsheet&quot; xmlns=&quot;urn:azure:m:v1&quot;&gt;&lt;r name=&quot;Spreadsheet&quot;&gt;&lt;e name=&quot;Endpoint1&quot; /&gt;&lt;/r&gt;&lt;/m&gt;" />
            </settings>
            <resourcereferences>
              <resourceReference name="DiagnosticStore" defaultAmount="[4096,4096,4096]" defaultSticky="true" kind="Directory" />
              <resourceReference name="EventStore" defaultAmount="[1000,1000,1000]" defaultSticky="false" kind="LogStore" />
            </resourcereferences>
          </role>
          <sCSPolicy>
            <sCSPolicyIDMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/SpreadsheetInstances" />
            <sCSPolicyUpdateDomainMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/SpreadsheetUpgradeDomains" />
            <sCSPolicyFaultDomainMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/SpreadsheetFaultDomains" />
          </sCSPolicy>
        </groupHascomponents>
      </components>
      <sCSPolicy>
        <sCSPolicyUpdateDomain name="SpreadsheetUpgradeDomains" defaultPolicy="[5,5,5]" />
        <sCSPolicyFaultDomain name="SpreadsheetFaultDomains" defaultPolicy="[2,2,2]" />
        <sCSPolicyID name="SpreadsheetInstances" defaultPolicy="[1,1,1]" />
      </sCSPolicy>
    </group>
  </groups>
  <implements>
    <implementation Id="afdc2dd5-9e03-48bd-a0c8-d2ff4e437682" ref="Microsoft.RedDog.Contract\ServiceContract\TemporaryDirectoryContract@ServiceDefinition">
      <interfacereferences>
        <interfaceReference Id="3b0d5c56-d25e-4f42-ba3f-e582ea4b0e39" ref="Microsoft.RedDog.Contract\Interface\Spreadsheet:Endpoint1@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/Spreadsheet:Endpoint1" />
          </inPort>
        </interfaceReference>
      </interfacereferences>
    </implementation>
  </implements>
</serviceModel>