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
// <copyright file="Errors.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.IO;
using System.Collections.Generic;
using System.Net;
using System.Collections.Specialized;
using System.Runtime.Serialization;
using System.Security.Permissions;
using System.Diagnostics.CodeAnalysis;
using System.Web;
using System.Linq;
using System.Xml;
using System.Xml.Linq;


// disable the generation of warnings for missing documentation elements for 
// public classes/members in this file
// justification is that this file contains many public constants whose names 
// sufficiently reflect their intended usage 
#pragma warning disable 1591

namespace Microsoft.Samples.ServiceHosting.StorageClient
{
    /// <summary>
    /// Error codes that can be returned by the storage service or the client library.
    /// These are divided into server errors and client errors depending on which side
    /// the error can be attributed to.
    /// </summary>
    public enum StorageErrorCode
    {
        None = 0,

        //Server errors
        ServiceInternalError = 1,
        ServiceTimeout,
        ServiceIntegrityCheckFailed,
        TransportError,
        ServiceBadResponse,

        //Client errors
        ResourceNotFound,
        AccountNotFound,
        ContainerNotFound,
        BlobNotFound,
        AuthenticationFailure,
        AccessDenied,
        ResourceAlreadyExists,
        ContainerAlreadyExists,
        BlobAlreadyExists,
        BadRequest,
        ConditionFailed,
        BadGateway
    }

    [Serializable]
    public class StorageExtendedErrorInformation
    {
        public string ErrorCode { get; internal set; }
        public string ErrorMessage { get; internal set; }
        public NameValueCollection AdditionalDetails { get; internal set; }
    }

    /// <summary>
    /// The base class for storage service exceptions
    /// </summary>
    [Serializable]
    public abstract class StorageException : Exception
    {
        /// <summary>
        /// The Http status code returned by the storage service
        /// </summary>
        public HttpStatusCode StatusCode { get; private set; }

        /// <summary>
        /// The specific error code returned by the storage service
        /// </summary>
        public StorageErrorCode ErrorCode { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        public StorageExtendedErrorInformation ExtendedErrorInformation { get; private set; }

        protected StorageException()
        {
        }

        protected StorageException(
            StorageErrorCode errorCode,
            string message,
            HttpStatusCode statusCode,
            StorageExtendedErrorInformation extendedErrorInfo,
            Exception innerException
            )
            : base(message, innerException)
        {
            this.ErrorCode = errorCode;
            this.StatusCode = statusCode;
            this.ExtendedErrorInformation = extendedErrorInfo;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="StorageException"/> class with
        /// serialized data.
        /// </summary>
        /// <param name="info">The <see cref="SerializationInfo"/> object that contains serialized object
        /// data about the exception being thrown</param>
        /// <param name="context">The <see cref="StreamingContext"/> object that contains contextual information
        /// about the source or destionation. </param>
        protected StorageException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
            if (null == info)
            {
                throw new ArgumentNullException("info");
            }

            this.StatusCode = (HttpStatusCode)info.GetValue("StatusCode", typeof(HttpStatusCode));
            this.ErrorCode = (StorageErrorCode)info.GetValue("ErrorCode", typeof(StorageErrorCode));
            this.ExtendedErrorInformation = (StorageExtendedErrorInformation)info.GetValue(
                        "ExtendedErrorInformation", typeof(StorageExtendedErrorInformation));
        }

        /// <summary>
        /// Sets the <see cref="SerializationInfo"/> object with additional exception information
        /// </summary>
        /// <param name="info">The <see cref="SerializationInfo"/> object that holds the 
        /// serialized object data.</param>
        /// <param name="context">The <see cref="StreamingContext"/> object that contains contextual information
        /// about the source or destionation. </param>
        [SecurityPermissionAttribute(SecurityAction.Demand, SerializationFormatter = true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (null == info)
            {
                throw new ArgumentNullException("info");
            }

            info.AddValue("StatusCode", this.StatusCode);
            info.AddValue("ErrorCode", this.ErrorCode);
            info.AddValue("ExtendedErrorInformation", this.ExtendedErrorInformation);
            base.GetObjectData(info, context);
        }

    }

    /// <summary>
    /// Server exceptions are those due to server side problems.
    /// These may be transient and requests resulting in such exceptions
    /// can be retried with the same parameters.
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1032:ImplementStandardExceptionConstructors",
        Justification = "Since this exception comes from the server, there must be an HTTP response code associated with it, hence we exclude the default constructor taking only a string but no status code.")]
    [Serializable]
    public class StorageServerException : StorageException
    {
        internal StorageServerException(
            StorageErrorCode errorCode,
            string message,
            HttpStatusCode statusCode,
            Exception innerException
            )
            : base(errorCode, message, statusCode, null, innerException)
        {
        }

        internal StorageServerException(
            StorageErrorCode errorCode,
            string message,
            HttpStatusCode statusCode,
            StorageExtendedErrorInformation extendedErrorInfo,
            Exception innerException
            )
            : base(errorCode, message, statusCode, extendedErrorInfo, innerException)
        {
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="StorageServerException"/> class with
        /// serialized data.
        /// </summary>
        /// <param name="info">The <see cref="SerializationInfo"/> object that contains serialized object
        /// data about the exception being thrown</param>
        /// <param name="context">The <see cref="StreamingContext"/> object that contains contextual information
        /// about the source or destionation. </param>
        protected StorageServerException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public StorageServerException()
        {
        }
    }

    /// <summary>
    /// Client side exceptions are due to incorrect parameters to the request.
    /// These requests should not be retried with the same parameters
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1032:ImplementStandardExceptionConstructors",
        Justification = "Since this exception comes from the server, there must be an HTTP response code associated with it, hence we exclude the default constructor taking only a string but no status code.")]
    [Serializable]
    public class StorageClientException : StorageException
    {
        internal StorageClientException(
            StorageErrorCode errorCode,
            string message,
            HttpStatusCode statusCode,
            StorageExtendedErrorInformation extendedErrorInfo,
            Exception innerException
            )
            : base(errorCode, message, statusCode, extendedErrorInfo, innerException)
        {
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="StorageClientException"/> class with
        /// serialized data.
        /// </summary>
        /// <param name="info">The <see cref="SerializationInfo"/> object that contains serialized object
        /// data about the exception being thrown</param>
        /// <param name="context">The <see cref="StreamingContext"/> object that contains contextual information
        /// about the source or destionation. </param>
        protected StorageClientException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public StorageClientException()
        {
        }

    }

    #region Error code strings that can be returned in the StorageExtendedErrorInformation.ErrorCode
    /// <summary>
    /// Error code strings that are common to all storage services
    /// </summary>
    public static class StorageErrorCodeStrings
    {
        public const string UnsupportedHttpVerb = "UnsupportedHttpVerb";
        public const string MissingContentLengthHeader = "MissingContentLengthHeader";
        public const string MissingRequiredHeader = "MissingRequiredHeader";
        public const string MissingRequiredXmlNode = "MissingRequiredXmlNode";
        public const string UnsupportedHeader = "UnsupportedHeader";
        public const string UnsupportedXmlNode = "UnsupportedXmlNode";
        public const string InvalidHeaderValue = "InvalidHeaderValue";
        public const string InvalidXmlNodeValue = "InvalidXmlNodeValue";
        public const string MissingRequiredQueryParameter = "MissingRequiredQueryParameter";
        public const string UnsupportedQueryParameter = "UnsupportedQueryParameter";
        public const string InvalidQueryParameterValue = "InvalidQueryParameterValue";
        public const string OutOfRangeQueryParameterValue = "OutOfRangeQueryParameterValue";
        public const string InvalidUri = "InvalidUri";
        public const string InvalidHttpVerb = "InvalidHttpVerb";
        public const string EmptyMetadataKey = "EmptyMetadataKey";
        public const string RequestBodyTooLarge = "RequestBodyTooLarge";
        public const string InvalidXmlDocument = "InvalidXmlDocument";
        public const string InternalError = "InternalError";
        public const string AuthenticationFailed = "AuthenticationFailed";
        public const string Md5Mismatch = "Md5Mismatch";
        public const string InvalidMd5 = "InvalidMd5";
        public const string OutOfRangeInput = "OutOfRangeInput";
        public const string InvalidInput = "InvalidInput";
        public const string OperationTimedOut = "OperationTimedOut";
        public const string ResourceNotFound = "ResourceNotFound";
        public const string InvalidMetadata = "InvalidMetadata";
        public const string MetadataTooLarge = "MetadataTooLarge";
        public const string ConditionNotMet = "ConditionNotMet";
        public const string InvalidRange = "InvalidRange";
        public const string ContainerNotFound = "ContainerNotFound";
        public const string ContainerAlreadyExists = "ContainerAlreadyExists";
        public const string ContainerDisabled = "ContainerDisabled";
        public const string ContainerBeingDeleted = "ContainerBeingDeleted";
        public const string ServerBusy = "ServerBusy";
    }

    /// <summary>
    /// Error code strings that are specific to blob service
    /// </summary>
    public static class BlobErrorCodeStrings
    {
        public const string InvalidBlockId = "InvalidBlockId";
        public const string BlobNotFound = "BlobNotFound";
        public const string BlobAlreadyExists = "BlobAlreadyExists";
        public const string InvalidBlobOrBlock = "InvalidBlobOrBlock";
        public const string InvalidBlockList = "InvalidBlockList";
    }

    /// <summary>
    /// Error code strings that are specific to queue service
    /// </summary>
    public static class QueueErrorCodeStrings
    {
        public const string QueueNotFound = "QueueNotFound";
        public const string QueueDisabled = "QueueDisabled";
        public const string QueueAlreadyExists = "QueueAlreadyExists";
        public const string QueueNotEmpty = "QueueNotEmpty";
        public const string QueueBeingDeleted = "QueueBeingDeleted";
        public const string PopReceiptMismatch = "PopReceiptMismatch";
        public const string InvalidParameter = "InvalidParameter";
        public const string MessageNotFound = "MessageNotFound";
        public const string MessageTooLarge = "MessageTooLarge";
        public const string InvalidMarker = "InvalidMarker";
    }

    /// <summary>
    /// Error code strings that are specific to queue service
    /// </summary>
    ///     public static class TableErrorCodeStrings
    public static class TableErrorCodeStrings
    {
        public const string XMethodNotUsingPost = "XMethodNotUsingPost";
        public const string XMethodIncorrectValue = "XMethodIncorrectValue";
        public const string XMethodIncorrectCount = "XMethodIncorrectCount";

        public const string TableHasNoProperties = "TableHasNoProperties";
        public const string DuplicatePropertiesSpecified = "DuplicatePropertiesSpecified";
        public const string TableHasNoSuchProperty = "TableHasNoSuchProperty";
        public const string DuplicateKeyPropertySpecified = "DuplicateKeyPropertySpecified";
        public const string TableAlreadyExists = "TableAlreadyExists";
        public const string TableNotFound = "TableNotFound";
        public const string EntityNotFound = "EntityNotFound";
        public const string EntityAlreadyExists = "EntityAlreadyExists";
        public const string PartitionKeyNotSpecified = "PartitionKeyNotSpecified";
        public const string OperatorInvalid = "OperatorInvalid";
        public const string UpdateConditionNotSatisfied = "UpdateConditionNotSatisfied";
        public const string PropertiesNeedValue = "PropertiesNeedValue";

        public const string PartitionKeyPropertyCannotBeUpdated = "PartitionKeyPropertyCannotBeUpdated";
        public const string TooManyProperties = "TooManyProperties";
        public const string EntityTooLarge = "EntityTooLarge";
        public const string PropertyValueTooLarge = "PropertyValueTooLarge";
        public const string InvalidValueType = "InvalidValueType";
        public const string TableBeingDeleted = "TableBeingDeleted";
        public const string TableServerOutOfMemory = "TableServerOutOfMemory";
        public const string PrimaryKeyPropertyIsInvalidType = "PrimaryKeyPropertyIsInvalidType";
        public const string PropertyNameTooLong = "PropertyNameTooLong";
        public const string PropertyNameInvalid = "PropertyNameInvalid";

        public const string BatchOperationNotSupported = "BatchOperationNotSupported";
        public const string JsonFormatNotSupported = "JsonFormatNotSupported";
        public const string MethodNotAllowed = "MethodNotAllowed";
        public const string NotImplemented = "NotImplemented";
    }
    #endregion

    #region Helper functions dealing with errors
    internal static partial class Utilities
    {
        internal static void ProcessUnexpectedStatusCode(HttpWebResponse response)
        {
            throw new StorageServerException(
                        StorageErrorCode.ServiceBadResponse,
                        response.StatusDescription,
                        response.StatusCode,
                        null
                        );
        }

        internal static Exception TranslateWebException(Exception e)
        {
            WebException we = e as WebException;
            if (null == we)
            {
                return e;
            }

            // If the response is not null, let's first see what the status code is.
            if (we.Response != null)
            {
                HttpWebResponse response = ((HttpWebResponse)we.Response);

                StorageExtendedErrorInformation extendedError =
                    GetExtendedErrorDetailsFromResponse(
                            response.GetResponseStream(),
                            response.ContentLength
                            );
                Exception translatedException = null;
                if (extendedError != null)
                {
                    translatedException = TranslateExtendedError(
                                                    extendedError,
                                                    response.StatusCode,
                                                    response.StatusDescription,
                                                    e);
                    if (translatedException != null)
                        return translatedException;
                }
                translatedException = TranslateFromHttpStatus(
                                            response.StatusCode,
                                            response.StatusDescription,
                                            extendedError,
                                            we
                                            );
                if (translatedException != null)
                    return translatedException;

            }

            switch (we.Status)
            {
                case WebExceptionStatus.RequestCanceled:
                    return new StorageServerException(
                        StorageErrorCode.ServiceTimeout,
                        "The server request did not complete within the specified timeout",
                        HttpStatusCode.GatewayTimeout,
                        we);

                case WebExceptionStatus.ConnectFailure:
                    return we;

                default:
                    return new StorageServerException(
                        StorageErrorCode.ServiceInternalError,
                        "The server encountered an unknown failure: " + e.Message,
                        HttpStatusCode.InternalServerError,
                        we
                        );
            }
        }

        internal static Exception TranslateFromHttpStatus(
                    HttpStatusCode statusCode,
                    string statusDescription,
                    StorageExtendedErrorInformation details,
                    Exception inner
                    )
        {
            switch (statusCode)
            {
                case HttpStatusCode.Forbidden:
                    return new StorageClientException(
                        StorageErrorCode.AccessDenied,
                        statusDescription,
                        HttpStatusCode.Forbidden,
                        details,
                        inner
                        );

                case HttpStatusCode.Gone:
                case HttpStatusCode.NotFound:
                    return new StorageClientException(
                        StorageErrorCode.ResourceNotFound,
                        statusDescription,
                        statusCode,
                        details,
                        inner
                        );

                case HttpStatusCode.BadRequest:
                    return new StorageClientException(
                        StorageErrorCode.BadRequest,
                        statusDescription,
                        statusCode,
                        details,
                        inner
                        );

                case HttpStatusCode.PreconditionFailed:
                case HttpStatusCode.NotModified:
                    return new StorageClientException(
                        StorageErrorCode.BadRequest,
                        statusDescription,
                        statusCode,
                        details,
                        inner);

                case HttpStatusCode.Conflict:
                    return new StorageClientException(
                        StorageErrorCode.ResourceAlreadyExists,
                        statusDescription,
                        statusCode,
                        details,
                        inner
                        );

                case HttpStatusCode.GatewayTimeout:
                    return new StorageServerException(
                        StorageErrorCode.ServiceTimeout,
                        statusDescription,
                        statusCode,
                        details,
                        inner
                        );

                case HttpStatusCode.RequestedRangeNotSatisfiable:
                    return new StorageClientException(
                        StorageErrorCode.BadRequest,
                        statusDescription,
                        statusCode,
                        details,
                        inner
                        );

                case HttpStatusCode.InternalServerError:
                    return new StorageServerException(
                        StorageErrorCode.ServiceInternalError,
                        statusDescription,
                        statusCode,
                        details,
                        inner
                        );

                case HttpStatusCode.BadGateway:
                    return new StorageServerException(
                        StorageErrorCode.BadGateway,
                        statusDescription,
                        statusCode,
                        details,
                        inner
                        );
            }
            return null;
        }

        private static Exception TranslateExtendedError(
                    StorageExtendedErrorInformation details,
                    HttpStatusCode statusCode,
                    string statusDescription,
                    Exception inner
                    )
        {
            StorageErrorCode errorCode = default(StorageErrorCode);
            switch (details.ErrorCode)
            {
                case StorageErrorCodeStrings.UnsupportedHttpVerb:
                case StorageErrorCodeStrings.MissingContentLengthHeader:
                case StorageErrorCodeStrings.MissingRequiredHeader:
                case StorageErrorCodeStrings.UnsupportedHeader:
                case StorageErrorCodeStrings.InvalidHeaderValue:
                case StorageErrorCodeStrings.MissingRequiredQueryParameter:
                case StorageErrorCodeStrings.UnsupportedQueryParameter:
                case StorageErrorCodeStrings.InvalidQueryParameterValue:
                case StorageErrorCodeStrings.OutOfRangeQueryParameterValue:
                case StorageErrorCodeStrings.InvalidUri:
                case StorageErrorCodeStrings.InvalidHttpVerb:
                case StorageErrorCodeStrings.EmptyMetadataKey:
                case StorageErrorCodeStrings.RequestBodyTooLarge:
                case StorageErrorCodeStrings.InvalidXmlDocument:
                case StorageErrorCodeStrings.InvalidXmlNodeValue:
                case StorageErrorCodeStrings.MissingRequiredXmlNode:
                case StorageErrorCodeStrings.InvalidMd5:
                case StorageErrorCodeStrings.OutOfRangeInput:
                case StorageErrorCodeStrings.InvalidInput:
                case StorageErrorCodeStrings.InvalidMetadata:
                case StorageErrorCodeStrings.MetadataTooLarge:
                case StorageErrorCodeStrings.InvalidRange:
                    errorCode = StorageErrorCode.BadRequest;
                    break;
                case StorageErrorCodeStrings.AuthenticationFailed:
                    errorCode = StorageErrorCode.AuthenticationFailure;
                    break;
                case StorageErrorCodeStrings.ResourceNotFound:
                    errorCode = StorageErrorCode.ResourceNotFound;
                    break;
                case StorageErrorCodeStrings.ConditionNotMet:
                    errorCode = StorageErrorCode.ConditionFailed;
                    break;
                case StorageErrorCodeStrings.ContainerAlreadyExists:
                    errorCode = StorageErrorCode.ContainerAlreadyExists;
                    break;
                case StorageErrorCodeStrings.ContainerNotFound:
                    errorCode = StorageErrorCode.ContainerNotFound;
                    break;
                case BlobErrorCodeStrings.BlobNotFound:
                    errorCode = StorageErrorCode.BlobNotFound;
                    break;
                case BlobErrorCodeStrings.BlobAlreadyExists:
                    errorCode = StorageErrorCode.BlobAlreadyExists;
                    break;
            }

            if (errorCode != default(StorageErrorCode))
                return new StorageClientException(
                                errorCode,
                                statusDescription,
                                statusCode,
                                details,
                                inner
                                );

            switch (details.ErrorCode)
            {
                case StorageErrorCodeStrings.InternalError:
                case StorageErrorCodeStrings.ServerBusy:
                    errorCode = StorageErrorCode.ServiceInternalError;
                    break;
                case StorageErrorCodeStrings.Md5Mismatch:
                    errorCode = StorageErrorCode.ServiceIntegrityCheckFailed;
                    break;
                case StorageErrorCodeStrings.OperationTimedOut:
                    errorCode = StorageErrorCode.ServiceTimeout;
                    break;
            }
            if (errorCode != default(StorageErrorCode))
                return new StorageServerException(
                                errorCode,
                                statusDescription,
                                statusCode,
                                details,
                                inner
                                );



            return null;
        }


        // This is the limit where we allow for the error message returned by the server.
        // Message longer than that will be truncated. 
        private const int ErrorTextSizeLimit = 8 * 1024;

        private static StorageExtendedErrorInformation GetExtendedErrorDetailsFromResponse(
            Stream httpResponseStream,
            long contentLength
            )
        {
            try
            {
                int bytesToRead = (int)Math.Max((long)contentLength, (long)ErrorTextSizeLimit);
                byte[] responseBuffer = new byte[bytesToRead];
                int bytesRead = CopyStreamToBuffer(httpResponseStream, responseBuffer, (int)bytesToRead);
                return GetErrorDetailsFromStream(
                            new MemoryStream(responseBuffer, 0, bytesRead, false)
                            );
            }
            catch (WebException)
            {
                //Ignore network errors when reading error details.
                return null;
            }
            catch (IOException)
            {
                return null;
            }
            catch (TimeoutException)
            {
                return null;
            }
        }

        private static StorageExtendedErrorInformation GetErrorDetailsFromStream(
            Stream inputStream
            )
        {
            StorageExtendedErrorInformation extendedError = new StorageExtendedErrorInformation();
            try
            {
                using (XmlReader reader = XmlReader.Create(inputStream))
                {
                    reader.Read();
                    reader.ReadStartElement(StorageHttpConstants.XmlElementNames.ErrorRootElement);
                    extendedError.ErrorCode = reader.ReadElementString(StorageHttpConstants.XmlElementNames.ErrorCode);
                    extendedError.ErrorMessage = reader.ReadElementString(StorageHttpConstants.XmlElementNames.ErrorMessage);
                    extendedError.AdditionalDetails = new NameValueCollection();

                    // After error code and message we can have a number of additional details optionally followed
                    // by ExceptionDetails element - we'll read all of these into the additionalDetails collection
                    do
                    {
                        if (reader.IsStartElement())
                        {
                            if (string.Compare(reader.LocalName, StorageHttpConstants.XmlElementNames.ErrorException, StringComparison.Ordinal) == 0)
                            {
                                // Need to read exception details - we have message and stack trace
                                reader.ReadStartElement(StorageHttpConstants.XmlElementNames.ErrorException);
                                extendedError.AdditionalDetails.Add(StorageHttpConstants.XmlElementNames.ErrorExceptionMessage,
                                    reader.ReadElementString(StorageHttpConstants.XmlElementNames.ErrorExceptionMessage));
                                extendedError.AdditionalDetails.Add(StorageHttpConstants.XmlElementNames.ErrorExceptionStackTrace,
                                    reader.ReadElementString(StorageHttpConstants.XmlElementNames.ErrorExceptionStackTrace));
                                reader.ReadEndElement();
                            }
                            else
                            {
                                string elementName = reader.LocalName;
                                extendedError.AdditionalDetails.Add(elementName, reader.ReadString());
                            }
                        }
                    }
                    while (reader.Read());
                }
            }
            catch (XmlException)
            {
                //If there is a parsing error we cannot return extended error information
                return null;
            }
            return extendedError;
        }

        internal static StorageExtendedErrorInformation GetExtendedErrorFromXmlMessage(string xmlErrorMessage)
        {
            string message = null;
            string errorCode = null;

            XName xnErrorCode = XName.Get(StorageHttpConstants.XmlElementNames.TableErrorCodeElement,
                StorageHttpConstants.XmlElementNames.DataWebMetadataNamespace);
            XName xnMessage = XName.Get(StorageHttpConstants.XmlElementNames.TableErrorMessageElement,
                StorageHttpConstants.XmlElementNames.DataWebMetadataNamespace);

            using (StringReader reader = new StringReader(xmlErrorMessage))
            {
                XDocument xDocument = null;
                try
                {
                    xDocument = XDocument.Load(reader);
                }
                catch (XmlException)
                {
                    // The XML could not be parsed. This could happen either because the connection 
                    // could not be made to the server, or if the response did not contain the
                    // error details (for example, if the response status code was neither a failure
                    // nor a success, but a 3XX code such as NotModified.
                    return null;
                }

                XElement errorCodeElement =
                    xDocument.Descendants(xnErrorCode).FirstOrDefault();

                if (errorCodeElement == null)
                    return null;

                errorCode = errorCodeElement.Value;

                XElement messageElement =
                    xDocument.Descendants(xnMessage).FirstOrDefault();

                if (messageElement != null)
                {
                    message = messageElement.Value;
                }

            }

            StorageExtendedErrorInformation errorDetails = new StorageExtendedErrorInformation();
            errorDetails.ErrorMessage = message;
            errorDetails.ErrorCode = errorCode;
            return errorDetails;
        }
    }


    #endregion
}