<?xml version="1.0" encoding="utf-8"?>
<serviceModel xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" name="TemporaryDirectory" generation="1" functional="0" release="0" Id="06d8a208-7637-4603-815d-bc7f0b6c2986" dslVersion="1.2.0.0" xmlns="http://schemas.microsoft.com/dsltools/RDSM">
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
    <implementation Id="af287620-e8fd-4b22-88dd-3b515a852f9d" ref="Microsoft.RedDog.Contract\ServiceContract\TemporaryDirectoryContract@ServiceDefinition">
      <interfacereferences>
        <interfaceReference Id="501a2a84-93de-4cea-8c87-8bba9dc61755" ref="Microsoft.RedDog.Contract\Interface\Spreadsheet:Endpoint1@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/TemporaryDirectory/TemporaryDirectoryGroup/Spreadsheet:Endpoint1" />
          </inPort>
        </interfaceReference>
      </interfacereferences>
    </implementation>
  </implements>
</serviceModel>