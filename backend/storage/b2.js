import {
  S3Client,
  PutObjectCommand,
  GetObjectCommand,
} from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";
import { randomUUID } from "crypto";

const CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function getConfig() {
  const keyId = process.env.B2_KEY_ID;
  const appKey = process.env.B2_APP_KEY || process.env.B2_APPLICATION_KEY;
  const bucket = process.env.B2_BUCKET || process.env.B2_BUCKET_NAME;
  const region = process.env.B2_REGION || "us-east-005";
  const endpoint =
    process.env.B2_ENDPOINT || `https://s3.${region}.backblazeb2.com`;
  const urlExpiry = Number(process.env.B2_URL_EXPIRY) || 600;

  if (!keyId || !appKey || !bucket) {
    throw new Error(
      "Backblaze B2 is not configured. Set B2_KEY_ID, B2_APP_KEY, and B2_BUCKET."
    );
  }

  return { keyId, appKey, bucket, region, endpoint, urlExpiry };
}

function getClient() {
  const { keyId, appKey, region, endpoint } = getConfig();

  return new S3Client({
    endpoint,
    region,
    forcePathStyle: true,
    credentials: {
      accessKeyId: keyId,
      secretAccessKey: appKey,
    },
  });
}

function objectKey(type, fileId) {
  return `${type}/${fileId}.xlsx`;
}

export async function uploadBuffer(buffer, { type, fileName }) {
  const { bucket, urlExpiry } = getConfig();
  const fileId = randomUUID();
  const key = objectKey(type, fileId);
  const client = getClient();

  await client.send(
    new PutObjectCommand({
      Bucket: bucket,
      Key: key,
      Body: buffer,
      ContentType: CONTENT_TYPE,
      Metadata: {
        originalname: fileName,
        type,
      },
    })
  );

  const downloadUrl = await getDownloadUrl(type, fileId, urlExpiry);

  return { fileId, fileName, downloadUrl, key };
}

export async function downloadFile(type, fileId) {
  const { bucket } = getConfig();
  const client = getClient();
  const key = objectKey(type, fileId);

  const response = await client.send(
    new GetObjectCommand({
      Bucket: bucket,
      Key: key,
    })
  );

  const chunks = [];
  for await (const chunk of response.Body) {
    chunks.push(chunk);
  }

  return {
    buffer: Buffer.concat(chunks),
    fileName: response.Metadata?.originalname || "workbook-output.xlsx",
  };
}

export async function getDownloadUrl(type, fileId, expiresIn) {
  const { bucket, urlExpiry } = getConfig();
  const client = getClient();
  const command = new GetObjectCommand({
    Bucket: bucket,
    Key: objectKey(type, fileId),
  });

  return getSignedUrl(client, command, {
    expiresIn: expiresIn ?? urlExpiry,
  });
}
