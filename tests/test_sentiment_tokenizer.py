"""Exercise the SentencePiece conversion used by the default sentiment model."""

from sentencepiece import sentencepiece_model_pb2
from tokenizers.models import BPE
from transformers.convert_slow_tokenizer import SentencePieceExtractor


def test_sentencepiece_model_converts_without_tiktoken(tmp_path):
    model = sentencepiece_model_pb2.ModelProto()
    model.trainer_spec.model_type = 2  # BPE, like the Cardiff sentiment model.
    model.trainer_spec.unk_id = 0
    for text in ("<unk>", "▁", "hello"):
        piece = model.pieces.add()
        piece.piece = text
        piece.score = 0

    path = tmp_path / "sentencepiece.bpe.model"
    path.write_bytes(model.SerializeToString())
    result = SentencePieceExtractor(str(path)).extract(BPE)

    assert result["vocab"]["hello"] == 2
