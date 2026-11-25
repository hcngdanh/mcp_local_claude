import { IMAGE_MIME_BY_EXT, IMAGE_RESIZE_STEPS, MAX_IMAGE_BLOCK_BYTES } from "../config.mjs";
import { log } from "../logger.mjs";
import { addContextNote, getExtension } from "./common.mjs";

export let sharpPromise = null;

export async function loadSharp() {
  if (!sharpPromise) {
    sharpPromise = import("sharp").then((module) => {
      const sharp = module.default || module.sharp || module;
      if (typeof sharp !== "function") {
        throw new Error("sharp loaded but did not expose a usable default export.");
      }
      return sharp;
    });
  }

  return sharpPromise;
}

export async function optimizeImageForBlock(buffer, mimeType) {
  const original = Buffer.from(buffer);

  if (original.length <= MAX_IMAGE_BLOCK_BYTES) {
    return {
      ok: true,
      buffer: original,
      mimeType,
      originalBytes: original.length,
      finalBytes: original.length,
      transformed: false,
    };
  }

  try {
    const sharp = await loadSharp();
    const metadata = await sharp(original, {
      animated: false,
      limitInputPixels: false,
    })
      .rotate()
      .metadata();

    const hasAlpha = !!metadata.hasAlpha;
    const outputMimeType = hasAlpha ? "image/webp" : "image/jpeg";

    for (const maxDimension of IMAGE_RESIZE_STEPS) {
      let pipeline = sharp(original, {
        animated: false,
        limitInputPixels: false,
      }).rotate();

      pipeline = pipeline.resize({
        width: maxDimension,
        height: maxDimension,
        fit: "inside",
        withoutEnlargement: true,
      });

      const compressed = hasAlpha
        ? await pipeline.webp({ quality: 76, effort: 4, alphaQuality: 80 }).toBuffer()
        : await pipeline.jpeg({ quality: 76, mozjpeg: true }).toBuffer();

      if (compressed.length <= MAX_IMAGE_BLOCK_BYTES) {
        return {
          ok: true,
          buffer: compressed,
          mimeType: outputMimeType,
          originalBytes: original.length,
          finalBytes: compressed.length,
          transformed: true,
        };
      }
    }

    const fallback = await sharp(original, {
      animated: false,
      limitInputPixels: false,
    })
      .rotate()
      .resize({
        width: 320,
        height: 320,
        fit: "inside",
        withoutEnlargement: true,
      })
      .webp({ quality: 60, effort: 4 })
      .toBuffer();

    if (fallback.length <= MAX_IMAGE_BLOCK_BYTES) {
      return {
        ok: true,
        buffer: fallback,
        mimeType: "image/webp",
        originalBytes: original.length,
        finalBytes: fallback.length,
        transformed: true,
      };
    }

    return {
      ok: false,
      reason: "single_image_too_large",
      originalBytes: original.length,
      finalBytes: fallback.length,
    };
  } catch (err) {
    log("Image optimization error", {
      mimeType,
      bytes: original.length,
      error: err?.message || String(err),
    });

    return {
      ok: false,
      reason: "image_optimization_failed",
      originalBytes: original.length,
      finalBytes: original.length,
    };
  }
}

export async function addImageBlock(context, buffer, mimeType) {
  if (!mimeType || !buffer) {
    return { attached: false, reason: "invalid_image_payload" };
  }

  if (context.imageBlocks.length >= context.maxImageBlocks) {
    addContextNote(
      context,
      `Some images were skipped because the maximum of ${context.maxImageBlocks} visual blocks was reached.`
    );
    return { attached: false, reason: "image_block_count_limit" };
  }

  const optimized = await optimizeImageForBlock(buffer, mimeType);
  if (!optimized.ok) {
    if (optimized.reason === "single_image_too_large") {
      addContextNote(
        context,
        `Some large images could not be attached because even after downscaling they exceeded the per-image MCP payload budget of ${Math.round(MAX_IMAGE_BLOCK_BYTES / 1024)} KB.`
      );
    } else {
      addContextNote(
        context,
        "Some images could not be optimized for inline visual preview and were skipped."
      );
    }

    return {
      attached: false,
      reason: optimized.reason,
      originalBytes: optimized.originalBytes,
      finalBytes: optimized.finalBytes,
    };
  }

  if (context.imageBytesUsed + optimized.finalBytes > context.maxTotalImageBytes) {
    addContextNote(
      context,
      `Some images were skipped because the total inline image budget of ${Math.round(context.maxTotalImageBytes / 1024)} KB was reached.`
    );
    return {
      attached: false,
      reason: "total_image_budget_reached",
      originalBytes: optimized.originalBytes,
      finalBytes: optimized.finalBytes,
    };
  }

  context.imageBlocks.push({
    type: "image",
    data: optimized.buffer.toString("base64"),
    mimeType: optimized.mimeType,
  });
  context.imageBytesUsed += optimized.finalBytes;

  if (optimized.transformed) {
    addContextNote(
      context,
      "Large images were downscaled and recompressed before being attached so they stay within Claude Desktop's MCP tool-result payload budget."
    );
  }

  return {
    attached: true,
    reason: optimized.transformed ? "compressed_for_delivery" : "attached_original",
    mimeType: optimized.mimeType,
    originalBytes: optimized.originalBytes,
    finalBytes: optimized.finalBytes,
    transformed: optimized.transformed,
  };
}

export function inferImageMimeType(ext, contentType = "") {
  const normalizedType = String(contentType || "").toLowerCase();
  if (normalizedType.startsWith("image/")) return normalizedType;
  return IMAGE_MIME_BY_EXT[String(ext || "").toLowerCase()] || null;
}

export function isImageAttachment(ext, contentType = "") {
  return !!inferImageMimeType(ext, contentType);
}

export async function collectZipImages(zip, mediaPrefix, context, maxImages) {
  let added = 0;

  for (const entry of zip.getEntries()) {
    if (entry.isDirectory || !entry.entryName.startsWith(mediaPrefix)) continue;
    if (added >= maxImages || context.imageBlocks.length >= context.maxImageBlocks) {
      break;
    }

    const ext = getExtension(entry.entryName);
    const mimeType = inferImageMimeType(ext);
    if (!mimeType) continue;

    const result = await addImageBlock(context, entry.getData(), mimeType);
    if (result.attached) {
      added += 1;
    }
  }

  return added;
}
