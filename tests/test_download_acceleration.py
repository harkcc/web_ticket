"""Regression contract for the browser Range download implementation."""

from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class DownloadAccelerationContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.frontend = (ROOT / "templates" / "index.html").read_text(encoding="utf-8")
        cls.backend = (ROOT / "web_ticket.py").read_text(encoding="utf-8")

    def test_frontend_probes_with_a_real_one_byte_range(self):
        self.assertIn("'Range': 'bytes=0-0'", self.frontend)
        self.assertIn("probeResp.status !== 206", self.frontend)
        self.assertNotIn("fetch(url, { method: 'HEAD' })", self.frontend)

    def test_frontend_rejects_malformed_or_wrong_sized_chunks(self):
        self.assertIn("parseContentRange", self.frontend)
        self.assertIn("contentRange.start !== chunk.start", self.frontend)
        self.assertIn("contentRange.total !== totalSize", self.frontend)
        self.assertIn("blob.size !== expectedSize", self.frontend)
        self.assertIn("mergedBlob.size !== totalSize", self.frontend)
        self.assertNotIn("resp.status !== 206 && resp.status !== 200", self.frontend)

    def test_dynamic_zip_probe_forces_safe_single_connection_download(self):
        self.assertIn("X-WebTicket-Range-Probe", self.frontend)
        self.assertIn("X-WebTicket-Download-Mode", self.frontend)
        self.assertIn("request.headers.get('X-WebTicket-Range-Probe') == '1'", self.backend)
        self.assertIn("response.headers['X-WebTicket-Download-Mode'] = 'single'", self.backend)

    def test_progress_updates_use_stable_element_ids(self):
        self.assertIn('id="mt-progress-size"', self.frontend)
        self.assertIn('id="mt-progress-percent"', self.frontend)
        self.assertNotIn("box.children[3].children[0].children[1]", self.frontend)


if __name__ == "__main__":
    unittest.main()
