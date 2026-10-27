# Limitations and data handling

- Word extraction is heuristic. The six synthetic fixtures demonstrate supported
  layouts, not arbitrary-document accuracy. Review generated facts and relationships.
- Literal-support validation rejects absent names, identifiers, amounts and years
  in LLM extraction. It does not establish that a fact belongs to the correct
  organization or validate free-form report prose. Full source-span provenance
  and an annotated extraction benchmark are future work.
- Input trees are limited to 32 levels and 1000 organization occurrences.
  Excel/Word files from untrusted sources should be processed in an isolated
  environment; comprehensive archive/resource limits are not implemented.
- Local mode avoids model requests. Model-assisted mode sends text to OLLAMA_URL.
  Cloud or LAN endpoints are a deliberate configuration choice.
- Anonymization masks supported fields and removes saved canvas state/geometry,
  including potentially embedded photos. It is not a certified anonymization
  system. Use the synthetic demo for public screenshots and shared artifacts.
- The editor risk score is a heuristic demonstration, not a legal or fraud finding.
- Automated checks validate extraction and output structure. Human visual checks
  in target versions of Microsoft Office are still needed for complex diagrams.
- The public demo uses fictitious names and placeholder identifiers. Do not infer
  valid government identifiers or actual contracts from those samples.

Saved canvas-only JSON restores the editor; use its toolbar for visual exports.
It does not contain a reliable semantic organization tree for Word/Excel conversion.
Anonymization of canvas-only states is rejected.
