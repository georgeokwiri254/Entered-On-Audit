# 🤖 Complete NER Training Guide for Reservation Email Extraction

This guide provides step-by-step instructions for training a DistilBERT NER model using your reservation email data.

## 📋 Overview

You now have a complete NER training pipeline with the following components:

- **Training Data Extractor**: Extracts data from MSG files using existing parsers
- **BIO Format Converter**: Converts parser outputs to NER training format
- **Google Colab Notebook**: Full training pipeline for DistilBERT
- **LabelStudio Setup**: Human annotation workflow for data quality
- **Validation Tools**: Data quality analysis and statistics

## 🚀 Step-by-Step Training Process

### Phase 1: Data Extraction (Local)

1. **Extract training data from MSG files:**
   ```bash
   python ner_training_data_extractor.py
   ```
   
   This will:
   - Process all MSG files in the Rules/ directory
   - Use existing parsers to extract MAIL_* fields
   - Generate training_data_YYYYMMDD_HHMMSS.json
   - Create extraction statistics and quality metrics

2. **Convert to BIO format:**
   ```bash
   python ner_bio_converter.py training_data_YYYYMMDD_HHMMSS.json
   ```
   
   This will:
   - Convert extracted fields to BIO token labels
   - Create train/validation/test splits (80/10/10)
   - Generate CoNLL and JSON formats
   - Save label mappings for the model

3. **Validate training data quality:**
   ```bash
   python ner_training_validator.py training_data_YYYYMMDD_HHMMSS.json --bio_data ner_bio_data/train_YYYYMMDD_HHMMSS.json
   ```
   
   This will:
   - Analyze data quality and coverage
   - Generate validation report and visualizations
   - Provide recommendations for improvement

### Phase 2: Data Annotation (Optional but Recommended)

4. **Set up LabelStudio for human annotation:**
   ```bash
   python labelstudio_setup.py
   ```
   
   Then:
   - Start LabelStudio: `./start_labelstudio.sh`
   - Access http://localhost:8080
   - Import training data for annotation
   - Correct extraction errors manually
   - Export corrected annotations

### Phase 3: Model Training (Google Colab)

5. **Upload to Google Colab:**
   - Open `DistilBERT_NER_Training_Colab.ipynb` in Google Colab
   - Enable GPU runtime (Runtime > Change runtime type > GPU)
   - Upload your training data files:
     - `train_YYYYMMDD_HHMMSS.json`
     - `val_YYYYMMDD_HHMMSS.json`
     - `test_YYYYMMDD_HHMMSS.json`
     - `label_mapping.json`

6. **Run training notebook:**
   - Execute all cells in sequence
   - Monitor training progress (typically 2-4 epochs)
   - Review evaluation metrics on test set
   - Download trained model

### Phase 4: Local Integration

7. **Integrate trained model:**
   ```python
   from transformers import pipeline
   
   # Load your trained model
   ner_model = pipeline("ner", 
                       model="./reservation-ner-model",
                       tokenizer="./reservation-ner-model",
                       aggregation_strategy="simple")
   
   # Extract entities
   entities = ner_model(email_text)
   ```

## 📊 Expected Results

Based on your current data (23 MSG files, 22 parsers), you can expect:

- **Training Dataset**: ~15-20 samples (small but focused)
- **Model Performance**: F1 score of 0.75-0.85 (with quality data)
- **Training Time**: 15-30 minutes on Google Colab GPU
- **Entity Coverage**: Strong performance on well-represented fields

## 🎯 Performance Optimization Tips

### Data Quality
- **High Priority Fields**: Focus on FIRST_NAME, ARRIVAL, DEPARTURE, C_T_S
- **Low Coverage Fields**: Improve parsers for NET_TOTAL, ADR, AMOUNT
- **Agency Balance**: Ensure representation from all major agencies

### Model Training
- **Batch Size**: Start with 16, reduce if GPU memory issues
- **Learning Rate**: 2e-5 works well for DistilBERT
- **Epochs**: 3-4 epochs usually sufficient
- **Early Stopping**: Monitor validation F1 score

### Production Integration
- **Confidence Thresholding**: Use 0.8+ for high confidence extractions
- **Fallback Strategy**: Keep existing parsers for low-confidence cases
- **Human Review**: Flag uncertain extractions for manual verification

## 🔧 Troubleshooting

### Common Issues

1. **Low Extraction Quality**
   ```bash
   # Check validation results
   python ner_training_validator.py training_data.json
   # Review parser outputs and improve regex patterns
   ```

2. **BIO Format Errors**
   ```bash
   # Validate BIO consistency
   python ner_bio_converter.py training_data.json
   # Check for token-label alignment issues
   ```

3. **Google Colab GPU Issues**
   - Use smaller batch size (8 or 16)
   - Enable FP16 mixed precision
   - Reduce sequence length if needed

4. **Poor Model Performance**
   - Increase training data volume
   - Improve label quality with LabelStudio
   - Add data augmentation techniques

### File Structure
```
Entered-On-Audit/
├── Rules/                          # MSG files and parsers
├── ner_training_data/              # Extracted training data
├── ner_bio_data/                   # BIO format data
├── ner_validation_results/         # Quality reports
├── labelstudio_data/               # Annotation data
├── ner_training_data_extractor.py  # Data extraction
├── ner_bio_converter.py            # BIO conversion
├── ner_training_validator.py       # Data validation
├── labelstudio_setup.py            # Annotation setup
└── DistilBERT_NER_Training_Colab.ipynb  # Training notebook
```

## 📈 Scaling Up

To improve model performance:

1. **More Data**: Collect additional MSG files from various agencies
2. **Data Augmentation**: Synthetic email generation with variations
3. **Active Learning**: Use model predictions to find annotation candidates
4. **Multi-task Learning**: Train on related tasks (date extraction, amount parsing)
5. **Ensemble Methods**: Combine NER model with existing rule-based parsers

## 🎉 Success Metrics

Your NER model is ready for production when:

- ✅ **F1 Score > 0.8** on test set
- ✅ **High confidence predictions** (>80% above 0.8 threshold)
- ✅ **Balanced performance** across all entity types
- ✅ **Consistent extractions** on new email formats
- ✅ **Better than existing parsers** on complex/edge cases

## 📞 Support

If you encounter issues:

1. Check the validation report for data quality recommendations
2. Review the generated statistics for insights
3. Use LabelStudio to manually correct problematic samples
4. Consider increasing training data if performance is low

The complete pipeline is designed to be iterative - run multiple training cycles with improved data quality for optimal results.

Good luck with your NER model training! 🚀