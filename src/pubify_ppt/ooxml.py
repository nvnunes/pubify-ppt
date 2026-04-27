from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import PurePosixPath
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
import posixpath


RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


@dataclass(frozen=True)
class MediaReplacement:
    """A direct media-part payload replacement for an OOXML package."""

    part_name: str
    blob: bytes


def patch_deck_package(
    source_path: object,
    output_path: object,
    deck: object,
    *,
    touched_slide_numbers: Iterable[int],
    media_replacements: Iterable[MediaReplacement] = (),
) -> None:
    """Write a deck by patching only changed OOXML package parts."""

    touched = tuple(sorted(set(touched_slide_numbers)))
    replacements = _replacement_parts(deck, touched)
    replacements.update(_media_replacements(media_replacements))
    with ZipFile(source_path, "r") as source_zip:
        source_names = set(source_zip.namelist())
        replacements.update(_new_or_changed_media_parts(deck, source_zip))
        replacements.update(_content_type_replacements(source_zip, replacements))
        final_relationships = _final_relationship_parts(source_zip, replacements)
        referenced_media = _referenced_media_parts(final_relationships)
        removed_media = {
            name
            for name in source_names
            if name.startswith("ppt/media/") and name not in referenced_media
        }
        _write_patched_zip(
            source_zip,
            output_path,
            replacements=replacements,
            removed_names=removed_media,
        )


def _replacement_parts(deck: object, touched_slide_numbers: tuple[int, ...]) -> dict[str, bytes]:
    replacements: dict[str, bytes] = {}
    for slide_number in touched_slide_numbers:
        slide = deck.slides[slide_number - 1]
        slide_name = _zip_name(str(slide.part.partname))
        replacements[slide_name] = slide.part.blob
        relationships = slide.part.rels.xml
        if relationships:
            replacements[_relationships_name(slide_name)] = relationships
    return replacements


def _media_replacements(media_replacements: Iterable[MediaReplacement]) -> dict[str, bytes]:
    return {
        _zip_name(replacement.part_name): replacement.blob
        for replacement in media_replacements
    }


def _new_or_changed_media_parts(deck: object, source_zip: ZipFile) -> dict[str, bytes]:
    replacements: dict[str, bytes] = {}
    source_names = set(source_zip.namelist())
    for part in deck.part.package.iter_parts():
        name = _zip_name(str(part.partname))
        if not name.startswith("ppt/media/"):
            continue
        blob = part.blob
        if name not in source_names or source_zip.read(name) != blob:
            replacements[name] = blob
    return replacements


def _content_type_replacements(source_zip: ZipFile, replacements: dict[str, bytes]) -> dict[str, bytes]:
    if not any(name.startswith("ppt/media/") and name.endswith(".png") for name in replacements):
        return {}
    content_types_name = "[Content_Types].xml"
    content_types = source_zip.read(content_types_name)
    root = ElementTree.fromstring(content_types)
    default_tag = f"{{{CONTENT_TYPES_NS}}}Default"
    has_png = any(
        child.tag == default_tag and child.attrib.get("Extension") == "png"
        for child in root
    )
    if has_png:
        return {}
    ElementTree.SubElement(
        root,
        default_tag,
        {"Extension": "png", "ContentType": "image/png"},
    )
    return {content_types_name: _xml_bytes(root)}


def _final_relationship_parts(source_zip: ZipFile, replacements: dict[str, bytes]) -> dict[str, bytes]:
    relationship_parts: dict[str, bytes] = {}
    for name in source_zip.namelist():
        if name.endswith(".rels"):
            relationship_parts[name] = replacements.get(name, source_zip.read(name))
    for name, blob in replacements.items():
        if name.endswith(".rels"):
            relationship_parts[name] = blob
    return relationship_parts


def _referenced_media_parts(relationship_parts: dict[str, bytes]) -> set[str]:
    referenced: set[str] = set()
    relationship_tag = f"{{{RELATIONSHIPS_NS}}}Relationship"
    for rels_name, blob in relationship_parts.items():
        if not blob:
            continue
        root = ElementTree.fromstring(blob)
        source_dir = _relationship_source_dir(rels_name)
        for relationship in root.findall(relationship_tag):
            target = relationship.attrib.get("Target")
            if not target or relationship.attrib.get("TargetMode") == "External":
                continue
            resolved = posixpath.normpath(posixpath.join(source_dir, target))
            if resolved.startswith("ppt/media/"):
                referenced.add(resolved)
    return referenced


def _write_patched_zip(
    source_zip: ZipFile,
    output_path: object,
    *,
    replacements: dict[str, bytes],
    removed_names: set[str],
) -> None:
    written: set[str] = set()
    with ZipFile(output_path, "w") as output_zip:
        for item in source_zip.infolist():
            if item.filename in removed_names:
                continue
            if item.filename in replacements:
                _write_zip_entry(output_zip, item, replacements[item.filename])
                written.add(item.filename)
            else:
                _write_zip_entry(output_zip, item, source_zip.read(item.filename))
                written.add(item.filename)
        for name, blob in sorted(replacements.items()):
            if name in written or name in removed_names:
                continue
            output_zip.writestr(_new_zip_info(name), blob)


def _write_zip_entry(output_zip: ZipFile, source_info: ZipInfo, blob: bytes) -> None:
    output_zip.writestr(source_info, blob)


def _new_zip_info(name: str) -> ZipInfo:
    info = ZipInfo(name)
    info.compress_type = ZIP_DEFLATED
    return info


def _relationships_name(part_name: str) -> str:
    path = PurePosixPath(part_name)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def _relationship_source_dir(rels_name: str) -> str:
    path = PurePosixPath(rels_name)
    if path.parent.name != "_rels":
        return str(path.parent)
    return str(path.parent.parent)


def _zip_name(partname: str) -> str:
    return partname.removeprefix("/")


def _xml_bytes(root: ElementTree.Element) -> bytes:
    ElementTree.register_namespace("", CONTENT_TYPES_NS)
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)
