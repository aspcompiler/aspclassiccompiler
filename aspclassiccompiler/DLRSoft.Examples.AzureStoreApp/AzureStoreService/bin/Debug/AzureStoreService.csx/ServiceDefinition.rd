<?xml version="1.0" encoding="utf-8"?>
<serviceModel xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" name="AzureStoreService" generation="1" functional="0" release="0" Id="7060094d-cb3f-4519-aa19-f2e13c9325f4" dslVersion="1.2.0.0" xmlns="http://schemas.microsoft.com/dsltools/RDSM">
  <groups>
    <group name="AzureStoreServiceGroup" generation="1" functional="0" release="0">
      <componentports>
        <inPort name="MVCAzureStore:HttpIn" protocol="http">
          <inToChannel>
            <lBChannelMoniker name="/AzureStoreService/AzureStoreServiceGroup/LB:MVCAzureStore:HttpIn" />
          </inToChannel>
        </inPort>
      </componentports>
      <settings>
        <aCS name="MVCAzureStoreInstances" defaultValue="[1,1,1]">
          <maps>
            <mapMoniker name="/AzureStoreService/AzureStoreServiceGroup/MapMVCAzureStoreInstances" />
          </maps>
        </aCS>
        <aCS name="MVCAzureStore:AccountName" defaultValue="">
          <maps>
            <mapMoniker name="/AzureStoreService/AzureStoreServiceGroup/MapMVCAzureStore:AccountName" />
          </maps>
        </aCS>
        <aCS name="MVCAzureStore:AccountSharedKey" defaultValue="">
          <maps>
            <mapMoniker name="/AzureStoreService/AzureStoreServiceGroup/MapMVCAzureStore:AccountSharedKey" />
          </maps>
        </aCS>
        <aCS name="MVCAzureStore:BlobStorageEndpoint" defaultValue="">
          <maps>
            <mapMoniker name="/AzureStoreService/AzureStoreServiceGroup/MapMVCAzureStore:BlobStorageEndpoint" />
          </maps>
        </aCS>
        <aCS name="MVCAzureStore:TableStorageEndpoint" defaultValue="">
          <maps>
            <mapMoniker name="/AzureStoreService/AzureStoreServiceGroup/MapMVCAzureStore:TableStorageEndpoint" />
          </maps>
        </aCS>
        <aCS name="MVCAzureStore:allowInsecureRemoteEndpoints" defaultValue="">
          <maps>
            <mapMoniker name="/AzureStoreService/AzureStoreServiceGroup/MapMVCAzureStore:allowInsecureRemoteEndpoints" />
          </maps>
        </aCS>
      </settings>
      <channels>
        <lBChannel name="LB:MVCAzureStore:HttpIn">
          <toPorts>
            <inPortMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore/HttpIn" />
          </toPorts>
        </lBChannel>
      </channels>
      <maps>
        <map name="MapMVCAzureStoreInstances" kind="Identity">
          <setting>
            <sCSPolicyIDMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStoreInstances" />
          </setting>
        </map>
        <map name="MapMVCAzureStore:AccountName" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore/AccountName" />
          </setting>
        </map>
        <map name="MapMVCAzureStore:AccountSharedKey" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore/AccountSharedKey" />
          </setting>
        </map>
        <map name="MapMVCAzureStore:BlobStorageEndpoint" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore/BlobStorageEndpoint" />
          </setting>
        </map>
        <map name="MapMVCAzureStore:TableStorageEndpoint" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore/TableStorageEndpoint" />
          </setting>
        </map>
        <map name="MapMVCAzureStore:allowInsecureRemoteEndpoints" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore/allowInsecureRemoteEndpoints" />
          </setting>
        </map>
      </maps>
      <components>
        <groupHascomponents>
          <role name="MVCAzureStore" generation="1" functional="0" release="0" software="F:\projects\dotnet35\VBParser80\AzureStoreAsp\AzureStoreService\obj\Debug\MVCAzureStore\" entryPoint="ucruntime" parameters="Microsoft.ServiceHosting.ServiceRuntime.Internal.WebRoleMain" memIndex="1024" hostingEnvironment="frontendfulltrust">
            <componentports>
              <inPort name="HttpIn" protocol="http" />
            </componentports>
            <settings>
              <aCS name="AccountName" defaultValue="" />
              <aCS name="AccountSharedKey" defaultValue="" />
              <aCS name="BlobStorageEndpoint" defaultValue="" />
              <aCS name="TableStorageEndpoint" defaultValue="" />
              <aCS name="allowInsecureRemoteEndpoints" defaultValue="" />
            </settings>
            <resourcereferences>
              <resourceReference name="EventStore" defaultAmount="[1000,1000,1000]" defaultSticky="false" kind="LogStore" />
            </resourcereferences>
            <eventstreams>
              <eventStream name="Microsoft.ServiceHosting.ServiceRuntime.RoleManager.Critical" kind="Default" severity="Critical" signature="Basic_string" />
              <eventStream name="Microsoft.ServiceHosting.ServiceRuntime.RoleManager.Error" kind="Default" severity="Error" signature="Basic_string" />
              <eventStream name="Critical" kind="Default" severity="Critical" signature="Basic_string" />
              <eventStream name="Error" kind="Default" severity="Error" signature="Basic_string" />
              <eventStream name="Warning" kind="OnDemand" severity="Warning" signature="Basic_string" />
              <eventStream name="Information" kind="OnDemand" severity="Info" signature="Basic_string" />
              <eventStream name="Verbose" kind="OnDemand" severity="Verbose" signature="Basic_string" />
            </eventstreams>
          </role>
          <sCSPolicy>
            <sCSPolicyIDMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStoreInstances" />
          </sCSPolicy>
        </groupHascomponents>
      </components>
      <sCSPolicy>
        <sCSPolicyID name="MVCAzureStoreInstances" defaultPolicy="[1,1,1]" />
      </sCSPolicy>
    </group>
  </groups>
  <implements>
    <implementation Id="6d41db0f-a06f-4b92-8068-8f3f77593ad4" ref="Microsoft.RedDog.Contract\ServiceContract\AzureStoreServiceContract@ServiceDefinition">
      <interfacereferences>
        <interfaceReference Id="f81791fe-32db-441f-ad48-c6390850c3f5" ref="Microsoft.RedDog.Contract\Interface\MVCAzureStore:HttpIn@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore:HttpIn" />
          </inPort>
        </interfaceReference>
      </interfacereferences>
    </implementation>
  </implements>
</serviceModel>