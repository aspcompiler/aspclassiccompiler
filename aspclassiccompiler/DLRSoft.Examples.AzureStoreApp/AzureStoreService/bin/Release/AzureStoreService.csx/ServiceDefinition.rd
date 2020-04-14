<?xml version="1.0" encoding="utf-8"?>
<serviceModel xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" name="AzureStoreService" generation="1" functional="0" release="0" Id="e95067b2-e3b0-4731-a49c-a0a01cc07a57" dslVersion="1.2.0.0" xmlns="http://schemas.microsoft.com/dsltools/RDSM">
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
          <role name="MVCAzureStore" generation="1" functional="0" release="0" software="F:\projects\dotnet35\VBParser80\AzureStoreAsp\AzureStoreService\obj\Release\MVCAzureStore\" entryPoint="ucruntime" parameters="Microsoft.ServiceHosting.ServiceRuntime.Internal.WebRoleMain" memIndex="1024" hostingEnvironment="frontend">
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
    <implementation Id="cd9be2b3-3818-4489-b03d-658dcf6a3002" ref="Microsoft.RedDog.Contract\ServiceContract\AzureStoreServiceContract@ServiceDefinition">
      <interfacereferences>
        <interfaceReference Id="2a39562e-75d8-4ebc-9574-33473ec9818f" ref="Microsoft.RedDog.Contract\Interface\MVCAzureStore:HttpIn@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/AzureStoreService/AzureStoreServiceGroup/MVCAzureStore:HttpIn" />
          </inPort>
        </interfaceReference>
      </interfacereferences>
    </implementation>
  </implements>
</serviceModel>