// ----------------------------------------------------------------------------------
// Microsoft Developer & Platform Evangelism
// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES 
// OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
// ----------------------------------------------------------------------------------
// The example companies, organizations, products, domain names,
// e-mail addresses, logos, people, places, and events depicted
// herein are fictitious.  No association with any real company,
// organization, product, domain name, email address, logo, person,
// places, or events is intended or should be inferred.
// ----------------------------------------------------------------------------------

//
// <copyright file="TableStorage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
// This file contains helper classes for accessing the table storage service:
//      - base classes for tables and table entities (table rows) containing the necessary 
//        partition key and row key values
//      - methods for applying the table storage authentication scheme and for handling 
//        authentication in DataServiceContext objects
//      - helper methods for creating, listing, and managing tables
//      - a set of table storage constants
//      - convenience methods for dealing with continuation tokens and paging
//      - simple error handling helpers
//      - retry semantics for table storage requests
// 
// Examples of how to make use of the classes and methods in this file are available 
// in the simple sample application contained in this solution.  


using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.Data.Services.Common;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;

// disable the generation of warnings for missing documentation elements for 
// public classes/members in this file
#pragma warning disable 1591

namespace Microsoft.Samples.ServiceHosting.StorageClient
{

    /// <summary>
    /// Class representing some important table storage constants.
    /// </summary>
    public static class TableStorageConstants
    {
        /// <summary>
        /// The maximum size of strings per property/column is 64 kB (that is 32k characters.)
        /// Note: This constant is smaller for the development storage table service.
        /// </summary>
        public static readonly int MaxStringPropertySizeInBytes = 64 * 1024;


        /// <summary>
        /// One character in the standard UTF-16 character presentation is 2 bytes.
        /// Note: This constant is smaller for the development storage table service.
        /// </summary>
        public static readonly int MaxStringPropertySizeInChars = MaxStringPropertySizeInBytes / 2;

        /// <summary>
        /// We want to prevent users from the pitfall of mixing up Utc and local time.
        /// Because of this we add some time to the minimum supported datetime.
        /// As a result, there will be no error condition from the server even 
        /// if a user converts the minimum supported date time to a local time and 
        /// stores this in a DateTime field.
        /// The local development storage support the SQL range of dates which is narrower than the
        /// one for the table storage service and so we use that value here. 
        /// </summary>
        public static readonly DateTime MinSupportedDateTime = DateTime.FromFileTime(0).ToUniversalTime().AddYears(200);

        //You can use this if you are programming against the real table storage service only but then your
        //code will not work against the local development table storage.
        //public static readonly DateTime MinSupportedDateTime = DateTime.FromFileTime(0).ToUniversalTime().AddDays(7);

        /// <summary>
        /// Internal constant for querying tables.
        /// </summary>
        internal const string TablesName = "Tables";

        /// <summary>
        /// Internal constant for querying tables.
        /// </summary>
        internal const string TablesQuery = "/" + TablesName;
    }


    /// <summary>
    /// API entry point for using structured storage. The underlying usage pattern is designed to be 
    /// similar to the one used in blob and queue services in this library. 
    /// Users create a TableStorage object by calling the static Create() method passing account credential 
    /// information to this method. The TableStorage object can then be used to create, delete and list tables. 
    /// There are two methods to get DataServiceContext objects that conform to the appropriate security scheme. 
    /// The first way is to call the GetDataServiceContext() method on TableStorage objects. The naming is again 
    /// chosen to conform to the convention in the other APIs for blob and queue services in this library. 
    /// This class can also be used as an adapter pattern. I.e., DataServiceContext objects can be created 
    /// independnt from a TableStorage object. Calling the Attach() method will make sure that the appropriate 
    /// security signing is used on these objects. This design was chosen to support various usage patterns that 
    /// might become necessary for autogenerated code.
    /// </summary>
    public class TableStorage
    {

        /// <summary>
        /// The default retry policy
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes",
                          Justification = "RetryPolicy is a non-mutable type")]
        public static readonly RetryPolicy DefaultRetryPolicy = RetryPolicies.NoRetry;


        /// <summary>
        /// Creates a TableStorage service object. This object is the entry point into the table storage API.
        /// </summary>
        /// <param name="baseUri">The base URI of the table storage service.</param>
        /// <param name="usePathStyleUris">Type of URI scheme used.</param>
        /// <param name="accountName">The account name.</param>
        /// <param name="base64Key">Base64 encoded version of the key.</param>
        /// <returns></returns>
        public static TableStorage Create(Uri baseUri,
                                          bool? usePathStyleUris,
                                          string accountName,
                                          string base64Key
                                         )
        {
            //We create a StorageAccountInfo and then extract the properties of that object.
            //This is because the constructor of StorageAccountInfo does normalization of BaseUri.
            StorageAccountInfo info = new StorageAccountInfo(
                                            baseUri,
                                            usePathStyleUris,
                                            accountName,
                                            base64Key
                                            );

            return new TableStorage(info.BaseUri, info.UsePathStyleUris, info.AccountName, info.Base64Key);
        }

        /// <summary>
        /// Creates a TableStorage object.
        /// </summary>
        public static TableStorage Create(StorageAccountInfo info)
        {
            return new TableStorage(info.BaseUri, info.UsePathStyleUris, info.AccountName, info.Base64Key);
        }

        /// <summary>
        /// Infers a list of tables from a DataServiceContext-derived type and makes sure
        /// those tables exist in the given service. The table endpoint information is retrieved from the 
        /// standard configuration settings.
        /// </summary>
        /// <remarks>
        /// Tables are inferred by finding all the public properties of type IQueryable&lt;T&gt; in 
        /// the provided type, where T is a type with an ID (in the case of table storage, this means it either
        /// has a [DataServiceKey("PartitionKey", "RowKey")] attribute in the class, or derives from
        /// the TableStorageEntity class included in this sample library (which in turn has that attribute).
        /// </remarks>
        public static void CreateTablesFromModel(Type serviceContextType)
        {
            CreateTablesFromModel(serviceContextType, StorageAccountInfo.DefaultTableStorageEndpointConfigurationString);
        }

        /// <summary>
        /// Infers a list of tables from a DataServiceContext-derived type and makes sure
        /// those tables exist in the given service.        
        /// </summary>
        /// <param name="serviceContextType">The DataServiceContext type from which the tables are inferred.</param>
        /// <param name="endpointConfiguration">A configuration string that is used to determine the table storage endpoint.</param>
        public static void CreateTablesFromModel(Type serviceContextType, string endpointConfiguration)
        {
            StorageAccountInfo account = StorageAccountInfo.GetAccountInfoFromConfiguration(endpointConfiguration);
            CreateTablesFromModel(serviceContextType, account);
        }

        /// <summary>
        /// Infers a list of tables from a DataServiceContext-derived type and makes sure
        /// those tables exist in the given service.        
        /// </summary>
        /// <param name="serviceContextType">The type of the DataServiceContext.</param>
        /// <param name="account">An object containing information about the table storage endpoint to be used.</param>
        public static void CreateTablesFromModel(Type serviceContextType, StorageAccountInfo account)
        {
            TableStorage tableStorage = TableStorage.Create(account);
            foreach (string tableName in DataServiceUtilities.EnumerateEntitySetNames(serviceContextType))
            {
                tableStorage.TryCreateTable(tableName);
            }
        }

        /// <summary>
        /// Creates a DataServiceContext object that takes care of implementing the table storage signing process.
        /// </summary>
        public TableStorageDataServiceContext GetDataServiceContext()
        {
            ResourceUriComponents uriComponents = new ResourceUriComponents(_accountName);
            Uri uri = HttpRequestAccessor.ConstructResourceUri(_baseUri, uriComponents, _usePathStyleUris);
            TableStorageDataServiceContext svc = new TableStorageDataServiceContext(uri, _accountName, _base64Key);
            if (svc != null)
            {
                svc.RetryPolicy = this.RetryPolicy;
            }
            return svc;
        }


        /// <summary>
        /// If the adaptor pattern with Attach() shall be used, this function can be used to generate the 
        /// table service base Uri depending on the path style syntax.
        /// </summary>
        static public Uri GetServiceBaseUri(Uri baseUri, bool usePathStyleUris, string accountName)
        {
            ResourceUriComponents uriComponents = new ResourceUriComponents(accountName);
            Uri uri = HttpRequestAccessor.ConstructResourceUri(baseUri, uriComponents, usePathStyleUris);
            return uri;
        }

        /// <summary>
        /// If the adaptor pattern with Attach() shall be used, this function can be used to generate the 
        /// table service base Uri depending on the path style syntax.
        /// </summary>
        static public Uri GetServiceBaseUri(StorageAccountInfo account)
        {
            return GetServiceBaseUri(account.BaseUri, account.UsePathStyleUris, account.AccountName);
        }

        /// <summary>
        /// If DataServiceContext objects are created at different places, this method can be called to configure the 
        /// DataServiceContext object to implement the required security scheme.
        /// </summary>
        public void Attach(DataServiceContext svc)
        {
            // this is an explicit way of dealing with situations where Attach() is called on objects that already 
            // have the necessary events hooked up to deal with table storage
            // in the event Attach() is called multiple times on a normal DataServiceContext object, we make sure to 
            // not doing authentication twice in the sending event itself
            if (svc is TableStorageDataServiceContext)
            {
                throw new ArgumentException("Cannot attach to a TableStorageDataServiceContext object. " +
                                            "These objects already contain the functionality for accessing the table storage service.");
            }
            new ContextRef(this, svc);
        }

        internal IEnumerable<TableStorageTable> ListTableImpl(DataServiceQuery<TableStorageTable> query)
        {
            IEnumerable<TableStorageTable> ret = null;

            RetryPolicy(() =>
            {
                try
                {
                    ret = query.Execute();
                }
                catch (InvalidOperationException e)
                {
                    if (TableStorageHelpers.CanBeRetried(e))
                    {
                        throw new TableRetryWrapperException(e);
                    }
                    throw;
                }
            });
            return ret;
        }

        /// <summary>
        /// Lists all the tables under this service's URL
        /// </summary>
        public IEnumerable<string> ListTables()
        {
            DataServiceContext svc = GetDataServiceContext();
            string nextKey = null;
            DataServiceQuery<TableStorageTable> localQuery;
            IEnumerable<TableStorageTable> tmp;
            svc.MergeOption = MergeOption.NoTracking;
            IQueryable<TableStorageTable> query = from t in svc.CreateQuery<TableStorageTable>(TableStorageConstants.TablesName)
                                                  select t;
            // result chunking
            // if we would not do this, the default value of 1000 is used before query result pagination
            // occurs
            query = query.Take(StorageHttpConstants.ListingConstants.MaxTableListResults);


            DataServiceQuery<TableStorageTable> orig = query as DataServiceQuery<TableStorageTable>;
            try
            {
                tmp = ListTableImpl(orig);
            }
            catch (InvalidOperationException e)
            {
                HttpStatusCode status;
                if (TableStorageHelpers.EvaluateException(e, out status) && status == HttpStatusCode.NotFound)
                {
                    yield break;
                }
                throw;
            }
            if (tmp == null)
            {
                yield break;
            }
            foreach (TableStorageTable table in tmp)
            {
                yield return table.TableName;
            }

            QueryOperationResponse qor = tmp as QueryOperationResponse;
            qor.Headers.TryGetValue(StorageHttpConstants.HeaderNames.PrefixForTableContinuation +
                                    StorageHttpConstants.HeaderNames.NextTableName,
                                    out nextKey);

            while (nextKey != null)
            {
                localQuery = orig;
                localQuery = localQuery.AddQueryOption(StorageHttpConstants.HeaderNames.NextTableName, nextKey);
                try
                {
                    tmp = ListTableImpl(localQuery);
                }
                catch (InvalidOperationException e)
                {
                    HttpStatusCode status;
                    if (TableStorageHelpers.EvaluateException(e, out status) && status == HttpStatusCode.NotFound)
                    {
                        yield break;
                    }
                    throw;
                }
                if (tmp == null)
                {
                    yield break;
                }
                foreach (TableStorageTable table in tmp)
                {
                    yield return table.TableName;
                }
                qor = tmp as QueryOperationResponse;
                qor.Headers.TryGetValue(StorageHttpConstants.HeaderNames.PrefixForTableContinuation +
                                        StorageHttpConstants.HeaderNames.NextTableName,
                                        out nextKey);
            }

        }

        /// <summary>
        /// Creates a new table in the service
        /// </summary>
        /// <param name="tableName">The name of the table to be created</param>
        public void CreateTable(string tableName)
        {
            ParameterValidator.CheckStringParameter(tableName, false, "tableName");
            if (!Utilities.IsValidTableName(tableName))
            {
                throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "The specified table name \"{0}\" is not valid!" +
                                            "Please choose a name that conforms to the naming conventions for tables!", tableName));
            }
            RetryPolicy(() =>
            {
                try
                {
                    DataServiceContext svc = GetDataServiceContext();
                    svc.AddObject(TableStorageConstants.TablesName, new TableStorageTable() { TableName = tableName });
                    svc.SaveChanges();
                }
                // exceptions are DataServiceClientException, DataServiceQueryException and DataServiceRequestException
                // all of the above exceptions are InvalidOperationExceptions
                catch (InvalidOperationException e)
                {
                    HttpStatusCode status;
                    if (TableStorageHelpers.CanBeRetried(e, out status))
                    {
                        if (status == HttpStatusCode.Conflict)
                        {
                            // don't retry in this case as this is an expected exception if the table exists
                            // just return the exception
                            throw;
                        }
                        else
                        {
                            throw new TableRetryWrapperException(e);
                        }
                    }
                    throw;
                }
            });
        }

        /// <summary>
        /// Tries to create a table with the given name.
        /// The main difference to the CreateTable method is that this function first queries the 
        /// table storage service whether the table already exists, before it tries to actually create 
        /// the table. The reason is that this 
        /// is more lightweight for the table storage service than always trying to create a table that 
        /// does already exist. Furthermore, as we expect that applications don't really randomly create
        /// tables, the additional roundtrip that is required for creating the table is necessary only very
        /// rarely.
        /// </summary>
        /// <param name="tableName">The name of the table.</param>
        /// <returns>True if the operation was completed successfully. False if the table already exists.</returns>
        public bool TryCreateTable(string tableName)
        {
            ParameterValidator.CheckStringParameter(tableName, false, "tableName");
            if (DoesTableExist(tableName))
            {
                return false;
            }
            try
            {
                CreateTable(tableName);
                return true;
            }
            catch (InvalidOperationException e)
            {
                HttpStatusCode status;
                if (TableStorageHelpers.EvaluateException(e, out status) && status == HttpStatusCode.Conflict)
                {
                    return false;
                }
                throw;
            }
        }

        /// <summary>
        /// Checks whether a table with the same name already exists.
        /// </summary>
        /// <param name="tableName">The name of the table to check.</param>
        /// <returns>True iff the table already exists.</returns>
        public bool DoesTableExist(string tableName)
        {
            ParameterValidator.CheckStringParameter(tableName, false, "tableName");
            bool tableExists = false;

            RetryPolicy(() =>
            {
                try
                {
                    DataServiceContext svc = GetDataServiceContext();
                    svc.MergeOption = MergeOption.NoTracking;
                    IEnumerable<TableStorageTable> query = from t in svc.CreateQuery<TableStorageTable>(TableStorageConstants.TablesName)
                                                           where t.TableName == tableName
                                                           select t;
                    tableExists = false;
                    try
                    {
                        // the query contains the whole primary key
                        // thus, if the query succeeds we can be sure that the table exists
                        (query as DataServiceQuery<TableStorageTable>).Execute();
                        tableExists = true;
                    }
                    catch (DataServiceQueryException e)
                    {
                        HttpStatusCode s;
                        if (TableStorageHelpers.EvaluateException(e, out s) && s == HttpStatusCode.NotFound)
                        {
                            tableExists = false;
                        }
                        else
                        {
                            throw;
                        }
                    }
                    catch (NullReferenceException ne)
                    {
                        //This is a workaround for bug in DataServiceQuery<T>.Execute. It throws a
                        //NullReferenceException instead of a DataServiceRequestException when it
                        //cannot connect to the the server. This workaround will be removed when
                        //the fix for this bug is released.
                        throw new DataServiceRequestException("Unable to connect to server.", ne);
                    }
                }
                catch (InvalidOperationException e)
                {
                    HttpStatusCode status;
                    if (TableStorageHelpers.CanBeRetried(e, out status))
                    {
                        throw new TableRetryWrapperException(e);
                    }
                    throw;
                }
            });
            return tableExists;
        }

        /// <summary>
        /// Deletes a table from the service.
        /// </summary>
        /// <param name="tableName">The name of the table to be deleted</param>
        public void DeleteTable(string tableName)
        {
            ParameterValidator.CheckStringParameter(tableName, false, "tableName");

            RetryPolicy(() =>
            {
                try
                {
                    DataServiceContext svc = GetDataServiceContext();
                    TableStorageTable table = new TableStorageTable() { TableName = tableName };
                    svc.AttachTo(TableStorageConstants.TablesName, table);
                    svc.DeleteObject(table);
                    svc.SaveChanges();
                }
                catch (InvalidOperationException e)
                {
                    HttpStatusCode status;
                    if (TableStorageHelpers.CanBeRetried(e, out status))
                    {
                        // we do this even thouh NoContent is currently not in the exceptions that 
                        // are retried
                        if (status == HttpStatusCode.NoContent || status == HttpStatusCode.NotFound)
                        {
                            // don't retry in this case, just return the exception
                            throw;
                        }
                        else
                        {
                            throw new TableRetryWrapperException(e);
                        }
                    }
                    throw;
                }
            });
        }

        /// <summary>
        /// Tries to delete the table with the given name. 
        /// </summary>
        /// <param name="tableName">The name of the table to delete.</param>
        /// <returns>True if the table was successfully deleted. False if the table does not exists.</returns>
        public bool TryDeleteTable(string tableName)
        {
            ParameterValidator.CheckStringParameter(tableName, false, "tableName");
            try
            {
                DeleteTable(tableName);
                return true;
            }
            catch (InvalidOperationException e)
            {
                HttpStatusCode status;
                if (TableStorageHelpers.EvaluateException(e, out status) && status == HttpStatusCode.NotFound)
                {
                    return false;
                }
                throw;
            }
        }

        /// <summary>
        /// The retry policy used for retrying requests
        /// </summary>
        public RetryPolicy RetryPolicy
        {
            get;
            set;
        }

        /// <summary>
        /// The base URI of the table storage service
        /// </summary>
        public Uri BaseUri
        {
            get
            {
                return this._baseUri;
            }
        }

        /// <summary>
        /// The name of the storage account
        /// </summary>
        public string AccountName
        {
            get
            {
                return this._accountName;
            }
        }

        /// <summary>
        /// Indicates whether to use/generate path-style or host-style URIs
        /// </summary>
        public bool UsePathStyleUris
        {
            get
            {
                return this._usePathStyleUris;
            }
        }

        /// <summary>
        /// The base64 encoded version of the key.
        /// </summary>
        internal string Base64Key
        {
            get
            {
                return _base64Key;
            }
        }

        internal protected TableStorage(Uri baseUri,
                            bool? usePathStyleUris,
                            string accountName,
                            string base64Key
                            )
        {
            this._baseUri = baseUri;
            this._accountName = accountName;
            this._base64Key = base64Key;
            if (usePathStyleUris == null)
            {
                this._usePathStyleUris = Utilities.StringIsIPAddress(baseUri.Host);
            }
            else
            {
                this._usePathStyleUris = usePathStyleUris.Value;
            }
            RetryPolicy = DefaultRetryPolicy;
        }

        private Uri _baseUri;
        private bool _usePathStyleUris;
        private string _accountName;
        private string _base64Key;
    }



    public static class TableStorageHelpers
    {

        #region Error handling helpers

        /// <summary>
        /// Checks whether the exception is or contains a DataServiceClientException and extracts the 
        /// returned http status code and extended error information.
        /// </summary>
        /// <param name="exception">The exception from which to extract information</param>
        /// <param name="status">The Http status code for the exception</param>
        /// <param name="extendedErrorInfo">Extended error information including storage service specific
        /// error code and error message</param>
        /// <returns>True if the exception is or contains a DataServiceClientException.</returns>
        public static bool EvaluateException(
            Exception exception,
            out HttpStatusCode status,
            out StorageExtendedErrorInformation extendedErrorInfo
            )
        {
            return EvaluateExceptionImpl(exception, out status, out extendedErrorInfo, true);
        }

        /// <summary>
        /// Checks whether the exception is or contains a DataServiceClientException and extracts the 
        /// returned http status code.
        /// </summary>
        /// <param name="exception">The exception from which to extract information</param>
        /// <param name="status">The Http status code for the exception</param>
        /// <returns>True if the exception is or contains a DataServiceClientException.</returns>
        public static bool EvaluateException(
            Exception exception,
            out HttpStatusCode status
            )
        {
            StorageExtendedErrorInformation extendedErrorInfo;
            return EvaluateExceptionImpl(exception, out status, out extendedErrorInfo, false);
        }

        private static bool EvaluateExceptionImpl(
            Exception e,
            out HttpStatusCode status,
            out StorageExtendedErrorInformation extendedErrorInfo,
            bool getExtendedErrors
            )
        {
            status = HttpStatusCode.Unused;
            extendedErrorInfo = null;
            while (e.InnerException != null)
            {
                e = e.InnerException;

                DataServiceClientException dsce = e as DataServiceClientException;
                if (dsce != null)
                {
                    status = (HttpStatusCode)dsce.StatusCode;
                    if (getExtendedErrors)
                        extendedErrorInfo = Utilities.GetExtendedErrorFromXmlMessage(dsce.Message);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Checks whether the exception is either a DataServiceClientException, a DataServiceQueryException or a 
        /// DataServiceRequestException.
        /// </summary>
        public static bool IsTableStorageException(Exception exception)
        {
            return ((exception is DataServiceClientException) || (exception is DataServiceQueryException) || (exception is DataServiceRequestException));
        }

        /// <summary>
        /// Only certain classes of errors should be retried. This method evaluates an exception 
        /// and returns whether this class of exception can be retried.
        /// </summary>
        /// <param name="e">The exception to analyze.</param>
        /// <param name="statusCode">The HttpStatusCode retrieved from the exception.</param>
        internal static bool CanBeRetried(InvalidOperationException e, out HttpStatusCode statusCode)
        {
            HttpStatusCode status;

            statusCode = HttpStatusCode.Unused;
            if (EvaluateException(e, out status))
            {
                statusCode = status;
                if (status == HttpStatusCode.RequestTimeout ||
                    // server error codes above 500
                    status == HttpStatusCode.ServiceUnavailable ||
                    status == HttpStatusCode.InternalServerError ||
                    status == HttpStatusCode.BadGateway ||
                    status == HttpStatusCode.GatewayTimeout)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Overload that does not retrun the HttpStatusCode.
        /// </summary>
        internal static bool CanBeRetried(InvalidOperationException e)
        {
            HttpStatusCode ignored;
            return CanBeRetried(e, out ignored);
        }

        #endregion

        #region Methods for checking properties to be inserted into a table

        /// <summary>
        /// Checks whether the string can be inserted in a table storage table. Throws an exception if 
        /// this is not the case.
        /// </summary>
        /// <param name="propertyValue"></param>
        public static void CheckStringProperty(string propertyValue)
        {
            if (string.IsNullOrEmpty(propertyValue))
            {
                throw new ArgumentException("The string cannot be null or empty!");
            }
            if (propertyValue.Length > TableStorageConstants.MaxStringPropertySizeInChars)
            {
                throw new ArgumentException("The string cannot be longer than the maximum string property size.");
            }
        }

        /// <summary>
        /// Checks whether the string can be inserted into a table storage table.
        /// </summary>
        public static bool ValidateStringProperty(string propertyValue)
        {
            if (string.IsNullOrEmpty(propertyValue))
            {
                return false;
            }
            if (propertyValue.Length > TableStorageConstants.MaxStringPropertySizeInChars)
            {
                return false;
            }
            return true;
        }

        #endregion

    }

    [DataServiceKey("TableName")]
    public class TableStorageTable
    {
        private string _tableName;


        /// <summary>
        /// The table name.
        /// </summary>
        public string TableName
        {
            get
            {
                return this._tableName;
            }

            set
            {
                ParameterValidator.CheckStringParameter(value, false, "TableName");
                this._tableName = value;
            }
        }

        public TableStorageTable()
        {
        }

        /// <summary>
        /// Creates a table with the specified name.
        /// </summary>
        /// <param name="name">The name of the table.</param>
        public TableStorageTable(string name)
        {
            ParameterValidator.CheckStringParameter(name, false, "name");
            this.TableName = name;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            TableStorageTable rhs = obj as TableStorageTable;

            if (rhs == null)
            {
                return false;
            }

            return (this.TableName == rhs.TableName);
        }

        public override int GetHashCode()
        {
            return this.TableName.GetHashCode();
        }
    }


    /// <summary>
    /// This class represents an entity (row) in a table in table storage.
    /// </summary>
    [CLSCompliant(false)]
    [DataServiceKey("PartitionKey", "RowKey")]
    public abstract class TableStorageEntity
    {

        public DateTime Timestamp
        {
            get;
            set;
        }

        /// <summary>
        /// The partition key of a table entity. The concatenation of the partition key 
        /// and row key must be unique per table.
        /// </summary>
        public virtual string PartitionKey
        {
            get;
            set;
        }

        /// <summary>
        /// The row key of a table entity.
        /// </summary>
        public virtual string RowKey
        {
            get;
            set;
        }

        /// <summary>
        /// Creates a TableStorageEntity object.
        /// </summary>
        protected TableStorageEntity(string partitionKey, string rowKey)
        {
            PartitionKey = partitionKey;
            RowKey = rowKey;
        }

        /// <summary>
        /// Creates a TableStorageEntity object.
        /// </summary>
        protected TableStorageEntity()
        {
        }

        
        /// <summary>
        /// Compares to entities.
        /// </summary>
        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            TableStorageEntity rhs = obj as TableStorageEntity;

            if (rhs == null)
            {
                return false;
            }

            return (this.PartitionKey == rhs.PartitionKey
                       && this.RowKey == rhs.RowKey);
        }
        

        /// <summary>
        /// Computes a HashCode for this object.
        /// </summary>
        public override int GetHashCode()
        {
            if (PartitionKey == null)
            {
                return base.GetHashCode();
            }
            if (!String.IsNullOrEmpty(this.RowKey))
            {
                return this.PartitionKey.GetHashCode() ^ this.RowKey.GetHashCode();
            }
            else
            {
                return this.PartitionKey.GetHashCode();
            }
        }
    }

    /// <summary>
    /// This class can be used for handling continuation tokens in TableStorage.
    /// </summary>
    /// <typeparam name="TElement"></typeparam>
    public class TableStorageDataServiceQuery<TElement>
    {

        private DataServiceQuery<TElement> _query;

        /// <summary>
        /// Objects of this class can be created using this constructor directly or by 
        /// calling a factory method on the TableStorageDataServiceContext class
        /// </summary>
        public TableStorageDataServiceQuery(DataServiceQuery<TElement> query)
            : this(query, RetryPolicies.NoRetry)
        {
        }

        /// <summary>
        /// Objects of this class can be created using this constructor directly or by 
        /// calling a factory method on the TableStorageDataServiceContext class
        /// </summary>
        public TableStorageDataServiceQuery(DataServiceQuery<TElement> query, RetryPolicy policy)
        {
            if (query == null)
            {
                throw new ArgumentNullException("query");
            }
            _query = query;
            RetryPolicy = policy;
        }

        /// <summary>
        /// Gets the underlying normal query object.
        /// </summary>
        public DataServiceQuery<TElement> Query
        {
            get
            {
                return _query;
            }
            set
            {
                _query = value;
            }
        }


        /// <summary>
        /// The retry policy used for retrying requests
        /// </summary>
        public RetryPolicy RetryPolicy
        {
            get;
            set;
        }

        /// <summary>
        /// Normal Execute() on the query without retry. Just maps to _query.Execute().
        /// </summary>
        /// <returns></returns>
        public IEnumerable<TElement> Execute()
        {
            return ExecuteWithRetries(RetryPolicies.NoRetry);
        }

        /// <summary>
        /// Calling Execute() on the query with the current retry policy.
        /// </summary>
        /// <returns>An IEnumerable respresenting the results of the query.</returns>
        public IEnumerable<TElement> ExecuteWithRetries()
        {
            return ExecuteWithRetries(RetryPolicy);
        }

        /// <summary>
        /// Calling Execute() on the query with the current retry policy.
        /// </summary>
        /// <param name="retry">The retry policy to be used for this request.</param>
        /// <returns>An IEnumerable respresenting the results of the query.</returns>
        public IEnumerable<TElement> ExecuteWithRetries(RetryPolicy retry)
        {
            IEnumerable<TElement> ret = null;
            if (retry == null)
            {
                throw new ArgumentNullException("retry");
            }
            retry(() =>
            {
                try
                {
                    ret = _query.Execute();
                }
                catch (InvalidOperationException e)
                {
                    if (TableStorageHelpers.CanBeRetried(e))
                    {
                        throw new TableRetryWrapperException(e);
                    }
                    throw;
                }
            });
            return ret;
        }

        /// <summary>
        /// Returns all results of the query and hides the complexity of continuation if 
        /// this is desired by a user. Users should be aware that this operation can return 
        /// many objects. Uses no retries.
        /// Important: this function does not call Execute immediately. Instead, it calls Execute() on 
        /// the query only when the result is enumerated. This is a difference to the normal 
        /// Execute() and Execute() with retry method.         
        /// </summary>
        /// <returns>An IEnumerable representing the results of the query.</returns>
        public IEnumerable<TElement> ExecuteAll()
        {
            return ExecuteAll(false);
        }

        /// <summary>
        /// Returns all results of the query and hides the complexity of continuation if 
        /// this is desired by a user. Users should be aware that this operation can return 
        /// many objects. This operation also uses the current retry policy.
        /// Important: this function does not call Execute immediately. Instead, it calls Execute() on 
        /// the query only when the result is enumerated. This is a difference to the normal 
        /// Execute() and Execute() with retry method. 
        /// </summary>
        /// <returns>An IEnumerable representing the results of the query.</returns>
        public IEnumerable<TElement> ExecuteAllWithRetries()
        {
            return ExecuteAll(true);
        }

        /// <summary>
        /// Returns all results of the query and hides the complexity of continuation if 
        /// this is desired by a user. Users should be aware that this operation can return 
        /// many objects. 
        /// Important: this function does not call Execute immediately. Instead, it calls Execute() on 
        /// the query only when the result is enumerated. This is a difference to the normal 
        /// Execute() and Execute() with retry method. 
        /// </summary>
        /// <param name="withRetry">Determines whether to use retries or not.</param>
        /// <returns>An IEnumerable representing the results of the query.</returns>
        public IEnumerable<TElement> ExecuteAll(bool withRetry)
        {
            IEnumerable<TElement> res;
            IEnumerable<TElement> tmp;
            string nextPartitionKey = null;
            string nextRowKey = null;
            DataServiceQuery<TElement> localQuery;

            if (_query == null)
            {
                throw new ArgumentException("The local DataServiceQuery element cannot be null!");
            }

            if (withRetry)
            {
                res = ExecuteWithRetries();
            }
            else
            {
                res = _query.Execute();
            }
            if (res == null)
            {
                yield break;
            }
            foreach (TElement item in res)
            {
                yield return item;
            }
            QueryOperationResponse qor = res as QueryOperationResponse;
            qor.Headers.TryGetValue(StorageHttpConstants.HeaderNames.PrefixForTableContinuation +
                                    StorageHttpConstants.HeaderNames.NextPartitionKey,
                                    out nextPartitionKey);
            qor.Headers.TryGetValue(StorageHttpConstants.HeaderNames.PrefixForTableContinuation +
                                    StorageHttpConstants.HeaderNames.NextRowKey,
                                    out nextRowKey);

            while (nextPartitionKey != null || nextRowKey != null)
            {
                localQuery = _query;
                if (nextPartitionKey != null)
                {
                    localQuery = localQuery.AddQueryOption(StorageHttpConstants.HeaderNames.NextPartitionKey, nextPartitionKey);
                }
                if (nextRowKey != null)
                {
                    localQuery = localQuery.AddQueryOption(StorageHttpConstants.HeaderNames.NextRowKey, nextRowKey);
                }
                if (withRetry)
                {
                    TableStorageDataServiceQuery<TElement> retryQuery = new TableStorageDataServiceQuery<TElement>(localQuery, RetryPolicy);
                    tmp = retryQuery.ExecuteWithRetries();
                }
                else
                {
                    tmp = localQuery.Execute();
                }
                if (tmp == null)
                {
                    yield break;
                }
                foreach (TElement item in tmp)
                {
                    yield return item;
                }
                qor = tmp as QueryOperationResponse;
                qor.Headers.TryGetValue(StorageHttpConstants.HeaderNames.PrefixForTableContinuation +
                                        StorageHttpConstants.HeaderNames.NextPartitionKey,
                                        out nextPartitionKey);
                qor.Headers.TryGetValue(StorageHttpConstants.HeaderNames.PrefixForTableContinuation +
                                        StorageHttpConstants.HeaderNames.NextRowKey,
                                        out nextRowKey);
            }
        }
    }

    /// <summary>
    /// The table storage-specific DataServiceContext class. It adds functionality for handling 
    /// the authentication process required by the table storage service.
    /// </summary>
    public class TableStorageDataServiceContext : DataServiceContext
    {
        private string _sharedKey;
        private string _accountName;

        /// <summary>
        /// Creates a DataServiceContext object and configures it so that it can be used with the table storage service.
        /// </summary>
        /// <param name="serviceRoot">The root URI of the service.</param>
        /// <param name="accountName">The account name.</param>
        /// <param name="sharedKey">The shared key associated with this service.</param>
        internal TableStorageDataServiceContext(Uri serviceRoot, string accountName, string sharedKey)
            : base(serviceRoot)
        {
            if (string.IsNullOrEmpty(accountName))
            {
                throw new ArgumentNullException("accountName");
            }
            if (string.IsNullOrEmpty(sharedKey))
            {
                throw new ArgumentNullException("sharedKey");
            }
            _sharedKey = sharedKey;
            _accountName = accountName;
            SendingRequest += new EventHandler<SendingRequestEventArgs>(DataContextSendingRequest);
            IgnoreMissingProperties = true;

            // we assume that this is the expected client behavior
            MergeOption = MergeOption.PreserveChanges;

            RetryPolicy = TableStorage.DefaultRetryPolicy;
        }

        /// <summary>
        /// Creates a DataServiceContext object and configures it so that it can be used with the table storage service.
        /// </summary>
        /// <param name="account">A StorageAccountInfo object containing information about how to access the table storage service.</param>
        public TableStorageDataServiceContext(StorageAccountInfo account)
            : this(TableStorage.GetServiceBaseUri(account), account.AccountName, account.Base64Key) { }


        /// <summary>
        /// Creates a DataServiceContext object and configures it so that it can be used with the table storage service.
        /// Information on the table storage endpoint is retrieved by accessing configuration settings in the app config section 
        /// of a Web.config or app config file, or by accessing settings in cscfg files.
        /// </summary>
        public TableStorageDataServiceContext()
            : this(StorageAccountInfo.GetAccountInfoFromConfiguration(StorageAccountInfo.DefaultTableStorageEndpointConfigurationString)) { }


        /// <summary>
        /// The retry policy used for retrying requests
        /// </summary>
        public RetryPolicy RetryPolicy
        {
            get;
            set;
        }

        /// <summary>
        /// Calls the SaveChanges() method and applies retry semantics.
        /// </summary>
        public DataServiceResponse SaveChangesWithRetries()
        {
            DataServiceResponse ret = null;
            RetryPolicy(() =>
            {
                try
                {
                    ret = SaveChanges();
                }
                catch (InvalidOperationException e)
                {
                    if (TableStorageHelpers.CanBeRetried(e))
                    {
                        throw new TableRetryWrapperException(e);
                    }
                    throw;
                }
            });
            return ret;
        }

        /// <summary>
        /// Calls the SaveChanges() method and applies retry semantics.
        /// </summary>
        public DataServiceResponse SaveChangesWithRetries(SaveChangesOptions options)
        {
            DataServiceResponse ret = null;
            RetryPolicy(() =>
            {
                try
                {
                    ret = SaveChanges(options);
                }
                catch (InvalidOperationException e)
                {
                    if (TableStorageHelpers.CanBeRetried(e))
                    {
                        throw new TableRetryWrapperException(e);
                    }
                    throw;
                }
            });
            return ret;
        }

        /// <summary>
        /// Callback method called whenever a request is sent to the table service. This 
        /// is where the signing of the request takes place.
        /// </summary>
        private void DataContextSendingRequest(object sender, SendingRequestEventArgs e)
        {
            HttpWebRequest request = e.Request as HttpWebRequest;

            // this setting can potentially result in very rare error conditions from HttpWebRequest
            // using retry policies, these error conditions are dealt with
            request.KeepAlive = true;

            // do the authentication
            byte[] key;
            SharedKeyCredentials credentials;

            Debug.Assert(_sharedKey != null);
            key = Convert.FromBase64String(_sharedKey);
            credentials = new SharedKeyCredentials(_accountName, key);
            credentials.SignRequestForSharedKeyLite(request, new ResourceUriComponents(_accountName, _accountName));
        }
    }

    #region Internal and private helpers

    /// <summary>
    /// Helper class to avoid long-lived references to context objects
    /// </summary>
    /// <remarks>
    /// Need to be careful not to maintain a reference to the context
    /// object from the auth adapter, since the adapter is probably
    /// long-lived and the context is not. This intermediate helper
    /// class is the one subscribing to context events, so when the
    /// context can be collected then this will be collectable as well.
    /// </remarks>
    internal class ContextRef
    {
        private TableStorage _adapter;

        public ContextRef(TableStorage adapter, DataServiceContext context)
        {
            if (adapter == null)
            {
                throw new ArgumentNullException("adapter");
            }
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }
            this._adapter = adapter;
            context.SendingRequest += this.HandleSendingRequest;
        }

        private void HandleSendingRequest(object sender, SendingRequestEventArgs e)
        {
            HttpWebRequest request = e.Request as HttpWebRequest;

            // first, we have to make sure that the request is not signed twice
            // this could happen if Attach() is called multiple times on the same DataServiceContext object
            WebHeaderCollection col = request.Headers;
            if (col != null)
            {
                foreach (string header in col)
                {
                    if (string.Compare(header, StorageHttpConstants.HeaderNames.StorageDateTime, StringComparison.Ordinal) == 0)
                    {
                        return;
                    }
                }
            }

            // this setting can potentially result in very rare error conditions from HttpWebRequest
            // using retry policies, these error conditions are dealt with
            request.KeepAlive = true;

            // do the authentication
            byte[] key;
            SharedKeyCredentials credentials;

            Debug.Assert(_adapter.Base64Key != null);
            key = Convert.FromBase64String(_adapter.Base64Key);
            credentials = new SharedKeyCredentials(_adapter.AccountName, key);
            credentials.SignRequestForSharedKeyLite(request, new ResourceUriComponents(_adapter.AccountName, _adapter.AccountName));
        }
    }

    /// <summary>
    /// The retry policies for blobs and queues deal with special StorageClient and StorageServer exceptions.
    /// In case of tables, we don't want to return these exceptions but instead the normal data service 
    /// exception. This class serves as a simple wrapper for these exceptions, and indicates that we 
    /// need retries.
    /// Data service exceptions are stored as inner exceptions.
    /// </summary>
    [Serializable]
    public class TableRetryWrapperException : Exception
    {
        /// <summary>
        /// Creates a TableRetryWrapperException object.
        /// </summary>
        public TableRetryWrapperException()
            : base()
        {
        }

        /// <summary>
        /// Creates a TableRetryWrapperException object.
        /// </summary>
        public TableRetryWrapperException(Exception innerException)
            : base(string.Empty, innerException)
        {
        }

        /// <summary>
        /// Creates a TableRetryWrapperException object.
        /// </summary>
        public TableRetryWrapperException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Creates a TableRetryWrapperException object.
        /// </summary>
        public TableRetryWrapperException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Creates a TableRetryWrapperException object.
        /// </summary>
        protected TableRetryWrapperException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

    }

    internal static class ParameterValidator
    {
        internal static void CheckStringParameter(string s, bool canBeNullOrEmpty, string name)
        {
            if (string.IsNullOrEmpty(s) && !canBeNullOrEmpty)
            {
                throw new ArgumentException(
                    string.Format(CultureInfo.InvariantCulture, "The parameter {0} cannot be null or empty.", name));
            }
            if (s.Length > TableStorageConstants.MaxStringPropertySizeInChars)
            {
                throw new ArgumentException(
                    string.Format(CultureInfo.InvariantCulture,
                        "The parameter {0} cannot be longer than {1} characters.",
                        name, TableStorageConstants.MaxStringPropertySizeInChars));
            }
        }
    }

    internal static class DataServiceUtilities
    {
        public static bool IsEntityType(Type t, Type contextType)
        {
            // ADO.NET data services convention: a type 't' is an entity if
            // 1) 't' has at least one key column
            // 2) there is a top level IQueryable<T> property in the context where T is 't' or a supertype of 't'
            // Non-primitive types that are not entity types become nested structures ("complex types" in EDM)

            if (!t.GetProperties().Any(p => IsKeyColumn(p))) return false;

            foreach (PropertyInfo pi in contextType.GetProperties())
            {
                if (typeof(IQueryable).IsAssignableFrom(pi.PropertyType))
                {
                    if (pi.PropertyType.GetGenericArguments()[0].IsAssignableFrom(t))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public static bool IsKeyColumn(PropertyInfo pi)
        {
            // Astoria convention:
            // 1) try the DataServiceKey attribute
            // 2) if not attribute, try <typename>ID
            // 3) finally, try just ID

            object[] attribs = pi.DeclaringType.GetCustomAttributes(typeof(DataServiceKeyAttribute), true);
            if (attribs != null && attribs.Length > 0)
            {
                Debug.Assert(attribs.Length == 1);
                return ((DataServiceKeyAttribute)attribs[0]).KeyNames.Contains(pi.Name);
            }

            if (pi.Name.Equals(pi.DeclaringType.Name + "ID", System.StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (pi.Name == "ID")
            {
                return true;
            }

            return false;
        }

        public static IEnumerable<PropertyInfo> EnumerateEntitySetProperties(Type contextType)
        {
            foreach (PropertyInfo prop in contextType.GetProperties())
            {
                if (typeof(IQueryable).IsAssignableFrom(prop.PropertyType) &&
                    prop.PropertyType.GetGenericArguments().Length > 0 &&
                    DataServiceUtilities.IsEntityType(prop.PropertyType.GetGenericArguments()[0], contextType))
                {
                    yield return prop;
                }
            }
        }

        public static IEnumerable<string> EnumerateEntitySetNames(Type contextType)
        {
            return EnumerateEntitySetProperties(contextType).Select(p => p.Name);
        }

    }

    #endregion
}