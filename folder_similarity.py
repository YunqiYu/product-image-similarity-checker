"""本地兼容入口：保留原命令，实际调用包内的相似度排序核心。"""

from product_image_pipeline.similarity_ranker import main


if __name__ == "__main__":
    raise SystemExit(main())
