// GCP ML Certification Training Hub - Full Curriculum & Expanded Slide Content

// --- State Management ---
const state = {
    currentSection: 'dashboard',
    slides: {
        m1_fundamentals: { current: 0, total: 10 },
        m2_features: { current: 0, total: 10 },
        m3_development: { current: 0, total: 10 },
        m4_pipelines: { current: 0, total: 10 },
        m5_production: { current: 0, total: 10 },
        m6_genai: { current: 0, total: 10 }
    },
    quiz: {
        currentQuestion: 0,
        score: 0,
        selectedChoice: null,
        answered: false,
        timer: 0,
        timerInterval: null,
        userAnswers: []
    },
    progress: {
        m1: 0,
        m2: 0,
        m3: 0,
        m4: 0,
        m5: 0,
        m6: 0,
        quiz: 0
    }
};

// --- Slide Content Data (60 Detailed slides Mapped to GCP Skills Portal) ---
const slideData = {
    m1_fundamentals: [
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 1 of 10",
            title: "Framing ML Problems & Storage Selection",
            body: `Machine learning engineering begins with framing: matching business objectives to model inputs and outputs. On GCP, storage choice dictates training throughput:
            <ul>
                <li><strong>Cloud Storage (GCS):</strong> Standard object storage for raw images, unstructured text, audio, and TFRecords. Highly scalable; integrates with Vertex AI pipelines.</li>
                <li><strong>BigQuery:</strong> Highly optimized for structured tabular data. Ideal for massive datasets where analytical querying and feature pre-aggregations are required.</li>
                <li><strong>Cloud Bigtable:</strong> Low-latency wide-column database. Ideal for real-time user-feature states served to online models.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 2 of 10",
            title: "GCP Ingestion & Processing Architecture",
            body: `GCP uses specific pipelines depending on data velocity:
            <ul>
                <li><strong>Cloud Pub/Sub:</strong> Decouples ingestion streams, acting as an event buffer. Essential for sub-second real-time streaming data.</li>
                <li><strong>Cloud Dataflow:</strong> Serverless Apache Beam processing engine. Handles both Batch and Streaming preprocessing transforms.</li>
                <li><strong>Cloud Dataproc:</strong> Managed Spark/Hadoop clusters. Best for migrating existing Spark ETL pipelines to Google Cloud without refactoring.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 3 of 10",
            title: "Pre-trained APIs vs. AutoML vs. Custom Training",
            body: `Choosing the right training mechanism is a frequent exam topic:
            <ul>
                <li><strong>Pre-trained APIs:</strong> Vision, Speech-to-Text, Translation, Natural Language. Best when task has general domain data and requires zero custom training.</li>
                <li><strong>AutoML:</strong> Vertex AI AutoML automatically conducts architecture search, training, and evaluation. Best for tabular, image, or text data when developer resource is low and custom model weights aren't needed.</li>
                <li><strong>Custom Training:</strong> Write custom TensorFlow, PyTorch, or JAX models. Required when you need specialized architectures, unique loss functions, or high-performance distributed scaling.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 4 of 10",
            title: "BigQuery ML (BQML)",
            body: `BQML allows SQL developers to train and serve models directly in BigQuery, bypassing data export bottlenecks:
            <ul>
                <li><strong>NATIVE EXECUTIONS:</strong> Runs training on BigQuery distributed compute slots. Supports linear regression, logistic regression, K-means, matrix factorization, and XGBoost.</li>
                <li><strong>MODEL DEPLOYMENT:</strong> You can export BQML models directly to Vertex AI Model Registry for serving prediction endpoints.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 5 of 10",
            title: "Pre-trained API Specializations",
            body: `Understand the specific tasks solved by Google Cloud's pre-trained APIs:
            <ul>
                <li><strong>Vision API:</strong> Performs face detection, logo detection, optical character recognition (OCR), and explicit content detection.</li>
                <li><strong>Natural Language API:</strong> Extracts entities (places, people, events), performs sentiment analysis, and tokenizes syntax.</li>
                <li><strong>Document AI:</strong> Extracts structured tables, key-value pairs, and text from unstructured documents (invoices, receipts, tax forms).</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 6 of 10",
            title: "AutoML Tabular, Image, and Text Pipelines",
            body: `Vertex AI AutoML automates deep learning pipeline creation:
            <ul>
                <li><strong>Feature Engineering:</strong> Automatically handles missing values, normalizes numeric ranges, and generates feature crosses.</li>
                <li><strong>Model Selection & Ensembling:</strong> Evaluates multiple model architectures (such as Wide & Deep, ResNet, XGBoost) and ensembles the best performers.</li>
                <li><strong>Explainable AI:</strong> Generates feature attribution metrics (Shapley values) automatically for tabular models.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 7 of 10",
            title: "BigQuery ML Advanced Features",
            body: `BQML supports advanced model structures:
            <ul>
                <li><strong>TRANSFORM Clause:</strong> Allows you to define feature preprocessing (like scaling and bucketization) inside the model definition. These transformations are packaged with the model, preventing skew.</li>
                <li><strong>Hyperparameter Tuning:</strong> Automatically runs hyperparameter searches (tuning variables like learning rate or tree depth) by setting options in the SQL statement.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 8 of 10",
            title: "Cloud Storage Integration Patterns",
            body: `Cloud Storage (GCS) is the primary data lake for training data:
            <ul>
                <li><strong>Storage Classes:</strong> Use Standard for active training datasets. Use Nearline/Coldline for archiving older models and datasets.</li>
                <li><strong>Bucket Notifications:</strong> Can be set to publish Pub/Sub events when new files arrive, triggering automated training pipelines.</li>
                <li><strong>IAM Controls:</strong> Use Service Accounts with <code>roles/storage.objectViewer</code> permissions to grant Vertex AI training containers access to datasets.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 9 of 10",
            title: "Cloud Dataproc for Managed Spark/Hadoop",
            body: `Dataproc is Google's managed Spark and Hadoop ecosystem:
            <ul>
                <li><strong>Migration Path:</strong> Best for enterprise migrations where legacy Spark SQL or Spark MLlib code needs to run in the cloud.</li>
                <li><strong>Ephemeral Clusters:</strong> Save cost by creating clusters programmatically, running Spark jobs, and immediately shutting them down.</li>
                <li><strong>Dataproc Serverless:</strong> Run Spark workloads directly without managing VM configurations or scaling policies.</li>
            </ul>`
        },
        {
            subtitle: "1. Big Data & ML Fundamentals • Slide 10 of 10",
            title: "Cloud Pub/Sub Streaming Ingestion Limits",
            body: `Pub/Sub is Google's globally scalable messaging queue:
            <ul>
                <li><strong>Retention Options:</strong> Messages are retained for up to 7 days by default.</li>
                <li><strong>Exactly-Once Delivery:</strong> Guarantees that messages are delivered exactly once to subscribers (such as Cloud Dataflow).</li>
                <li><strong>Dead-Letter Topics:</strong> Directs failed messages that cannot be processed to a separate topic for debugging and troubleshooting.</li>
            </ul>`
        }
    ],
    m2_features: [
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 1 of 10",
            title: "Feature Transformations & Scaling",
            body: `Raw data is rarely ready for ML training. Feature engineering transforms inputs to improve model convergence:
            <ul>
                <li><strong>Normalization (MinMax):</strong> Scales features to [0, 1]. Best when data distribution is bounded and does not follow Gaussian curves.</li>
                <li><strong>Standardization (Z-Score):</strong> Scales data to 0 mean and unit variance. Highly recommended for neural networks and algorithms that assume normal distributions.</li>
                <li><strong>Bucketization / Quantization:</strong> Converts continuous features to discrete bins (e.g., age ranges). Converts non-linear features into linear categorical inputs.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 2 of 10",
            title: "Feature Crosses & Vocabularies",
            body: `Capturing feature interactions:
            <ul>
                <li><strong>Feature Crosses:</strong> Mathematically combines two or more categorical features (e.g., <code>[latitude_bucket] × [longitude_bucket]</code>). This allows linear models to learn non-linear decision boundaries.</li>
                <li><strong>Vocabulary Mapping:</strong> Maps categorical text tokens to unique integers, followed by One-Hot encoding or dense trainable Embeddings.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 3 of 10",
            title: "Preprocessing Skew: Dataflow vs TF.Transform",
            body: `A critical exam theme is preventing **train-serve skew** during preprocessing:
            <ul>
                <li><strong>Dataflow Pipeline:</strong> Calculates parameters (like dataset Mean/Variance) and transforms training files. However, during online serving, client must manually code the same transformations. Any difference causes skew.</li>
                <li><strong>TF.Transform (TFX):</strong> Computes constants across the entire dataset (like dataset maximums) during training, and compiles them into a **Preprocessing Graph**. This preprocessing graph is exported *directly* with the TensorFlow model, ensuring matching logic.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 4 of 10",
            title: "Vertex AI Feature Store Concepts",
            body: `Centralizes feature registration, sharing, and low-latency serving:
            <ul>
                <li><strong>Data Model:</strong> Entity Types (e.g., <code>user_id</code>) container features (e.g., <code>clicks_last_hour</code>). Synchronizes directly via Feature Views.</li>
                <li><strong>Online Store serving:</strong> Backed by Google Cloud Bigtable or Redis for ultra-low latency inference fetches.</li>
                <li><strong>Offline Store serving:</strong> Extracts historical feature snapshots to BigQuery for training using point-in-time correctness joins to prevent leakage.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 5 of 10",
            title: "Dealing with Missing Data & Outliers",
            body: `Missing features and outliers can skew model predictions:
            <ul>
                <li><strong>Imputation:</strong> Fill missing numeric fields using Mean, Median, or Mode values. For categorical features, introduce a dedicated 'Unknown' category.</li>
                <li><strong>Feature Clipping:</strong> Cap outliers at specific percentiles (e.g., 1st and 99th percentiles) to prevent extreme values from distorting weights.</li>
                <li><strong>Log Transformations:</strong> Apply log scaling (e.g., <code>log(x + 1)</code>) to right-skewed features (such as price or traffic volumes) to compress distributions.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 6 of 10",
            title: "Representing Categorical Features",
            body: `Converting category text into numerical dimensions:
            <ul>
                <li><strong>One-Hot Encoding:</strong> Creates a binary column for every unique category. Best for low cardinality features (e.g., days of the week).</li>
                <li><strong>Hashing Trick:</strong> Hashes categories into a fixed number of bins. Handles open-ended vocabularies and high cardinality features, but introduces hash collisions.</li>
                <li><strong>Embeddings:</strong> Learns low-dimensional dense vectors representing categories. Captures semantic relationships between items.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 7 of 10",
            title: "Feature Selection & Dimensionality Reduction",
            body: `Reduce overfitting and training costs by removing redundant features:
            <ul>
                <li><strong>Principal Component Analysis (PCA):</strong> Unsupervised technique that projects features into orthogonal principal components, maximizing variance capture.</li>
                <li><strong>L1 Regularization (Lasso):</strong> Adds absolute coefficient penalties during training, driving non-essential feature weights strictly to zero.</li>
                <li><strong>Correlation Matrices:</strong> Identify and remove highly collinear features that destabilize model weights.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 8 of 10",
            title: "TFX TF.Transform Implementation",
            body: `Writing production TFX transformations:
            <ul>
                <li><strong>preprocessing_fn:</strong> The central user function containing feature manipulations.</li>
                <li><strong>tft utilities:</strong> Use TensorFlow Transform functions (like <code>tft.scale_to_z_score</code> or <code>tft.compute_and_apply_vocabulary</code>) to calculate statistics across the dataset.</li>
                <li><strong>Export logic:</strong> The generated transform output is saved as a SavedModel directory and loaded during serving.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 9 of 10",
            title: "Vertex AI Feature Store Online Serving",
            body: `High-performance serving for real-time predictions:
            <ul>
                <li><strong>Storage Layer:</strong> Uses Cloud Bigtable or a dedicated Redis cluster for sub-15ms feature lookups.</li>
                <li><strong>Read API:</strong> Endpoint applications call the <code>ReadFeatureValues</code> REST/gRPC API to fetch feature vectors.</li>
                <li><strong>Concurrency:</strong> Optimized for high-throughput serving, handling thousands of queries per second.</li>
            </ul>`
        },
        {
            subtitle: "2. Feature Engineering & Preproc. • Slide 10 of 10",
            title: "Vertex AI Feature Store Offline Serving",
            body: `Extracting historical features for training datasets:
            <ul>
                <li><strong>Batch Export:</strong> Runs point-in-time joins to compile feature values matching historical event timestamps.</li>
                <li><strong>Data Leakage Prevention:</strong> Guarantees that the training set only contains feature values that were active *before* the target prediction event.</li>
                <li><strong>Export Formats:</strong> Supports exporting features directly to BigQuery tables or CSV files in Cloud Storage.</li>
            </ul>`
        }
    ],
    m3_development: [
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 1 of 10",
            title: "Vertex AI Custom Training & Containers",
            body: `Vertex AI Custom Training runs code inside isolated Docker containers:
            <ul>
                <li><strong>Pre-built Containers:</strong> Google provides ready-to-use Docker images for TensorFlow, PyTorch, and Scikit-learn, containing CUDA drivers.</li>
                <li><strong>Custom Containers:</strong> Build a custom Dockerfile if you require specific library versions, C++ extensions, or non-standard model libraries.</li>
                <li><strong>Entrypoint rules:</strong> Custom containers must accept specific environment variables (like <code>AIP_MODEL_DIR</code>) to export model artifacts to GCS.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 2 of 10",
            title: "Distributed Training Frameworks",
            body: `When datasets exceed single-machine GPU memory, you must distribute training:
            <ul>
                <li><strong>Data Parallelism:</strong> Splits the batch of training records across workers. Each worker runs a copy of the model, updates local gradients, and aggregates weights (e.g., via <code>MultiWorkerMirroredStrategy</code>).</li>
                <li><strong>Model Parallelism:</strong> Splits the model architecture itself (e.g., layer divisions) across different GPU nodes because the model weights exceed single-device memory.</li>
                <li><strong>Vertex AI Reduction Server:</strong> Speeds up distributed training by aggregating gradients across compute worker instances, reducing networking latency.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 3 of 10",
            title: "Hardware Profiles: GPUs vs TPUs",
            body: `Choosing the right accelerator:
            <ul>
                <li><strong>GPUs (Nvidia A100/H100):</strong> Highly flexible. Best for PyTorch custom operations, computer vision networks (convolutions), and dynamic neural architectures.</li>
                <li><strong>TPUs (Tensor Processing Units):</strong> Custom ASICs built by Google. Optimized for large matrix multiplication (GEMM). Best for transformer networks, large-batch TensorFlow jobs, and LLM training.</li>
                <li><strong>TPU Rule:</strong> Requires static input shapes. Operations that rely heavily on custom C++ kernels or dynamic sizing should remain on GPUs.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 4 of 10",
            title: "Hyperparameter Tuning with Vizier",
            body: `Vertex AI Vizier is a managed optimization service for hyperparameter tuning:
            <ul>
                <li><strong>Bayesian Optimization:</strong> Instead of executing exhaustive Grid Search, Vizier builds a statistical surrogate model of loss and chooses parameter combinations that balance exploration and exploitation.</li>
                <li><strong>Vizier Study:</strong> Tracks parameter boundaries (e.g., learning rate log scale) and metrics to output trial metrics, finding the global optimum in fewer trials.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 5 of 10",
            title: "Distributed Training with Reduction Server",
            body: `Gradient aggregation in high-scale distributed training:
            <ul>
                <li><strong>AllReduce Bottleneck:</strong> Replicating gradients across hundreds of GPU workers leads to network congestion.</li>
                <li><strong>Reduction Server:</strong> Acts as a serverless parameter aggregator, collecting gradients from workers, performing reduction math, and returning updated weights.</li>
                <li><strong>Performance:</strong> Reduces training time for large models by up to 30% by decoupling network operations from worker computations.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 6 of 10",
            title: "TensorFlow Distributed Configurations",
            body: `Configuring TensorFlow distribution across nodes:
            <ul>
                <li><strong>TF_CONFIG Variable:</strong> A JSON environment variable that specifies the cluster layout, including IP addresses and roles.</li>
                <li><strong>Worker Roles:</strong> Defines 'chief' (coordinates training, saves checkpoints), 'worker' (runs training), and 'ps' (parameter servers that hold weights).</li>
                <li><strong>Automation:</strong> Vertex AI automatically generates and injects the correct TF_CONFIG variables into custom training containers.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 7 of 10",
            title: "PyTorch on Vertex AI (torchrun CLI)",
            body: `Orchestrating PyTorch custom training:
            <ul>
                <li><strong>torchrun:</strong> PyTorch elastic launcher CLI tool. Automates process spawning, rank assignments, and master IP lookups.</li>
                <li><strong>Dataset Loading:</strong> Use PyTorch DataLoader with <code>DistributedSampler</code> to split the training batches across worker nodes.</li>
                <li><strong>Example execution:</strong> Run <code>torchrun --nproc_per_node=X main.py</code> inside custom containers.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 8 of 10",
            title: "Model Evaluation Metrics",
            body: `Selecting appropriate evaluation metrics for model verification:
            <ul>
                <li><strong>ROC-AUC:</strong> Measures classification performance across all thresholds. Best for balanced datasets.</li>
                <li><strong>PR-AUC:</strong> Precision-Recall Area Under the Curve. Focuses on the positive class; highly recommended for imbalanced datasets (e.g. fraud detection).</li>
                <li><strong>Confusion Matrix:</strong> Tracks False Positives and False Negatives, assisting in optimization threshold selections.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 9 of 10",
            title: "Transfer Learning & Fine-tuning",
            body: `Leveraging pre-trained models for custom tasks:
            <ul>
                <li><strong>Frozen Layers:</strong> Keep the early layer weights of a pre-trained model locked to retain general feature representations (e.g. image edges).</li>
                <li><strong>Custom Head:</strong> Append new trainable dense layers to classify the model output for your target classes.</li>
                <li><strong>Learning Rate Schedule:</strong> Use low learning rates during fine-tuning to prevent destroying pre-trained weights.</li>
            </ul>`
        },
        {
            subtitle: "3. ML Model Dev & Scaling • Slide 10 of 10",
            title: "Vertex AI Vizier Early Stopping",
            body: `Vizier helps save compute resources during hyperparameter searches:
            <ul>
                <li><strong>Median Stopping Rule:</strong> Vizier evaluates trials mid-run. If a trial's performance is below the median of completed trials at the same step, it terminates the trial early.</li>
                <li><strong>Decay Optimization:</strong> Terminates trials that show stagnant metric improvements over multiple epochs.</li>
            </ul>`
        }
    ],
    m4_pipelines: [
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 1 of 10",
            title: "Why Orchestrate ML Workflows?",
            body: `Deploying an ML model is not a one-off task. A production model requires continuous pipeline executions to handle data updates. ML Orchestration ensures:
            <ul>
                <li><strong>Reproducibility:</strong> Every step (data ingestion, validation, training, evaluator, registry) runs in a defined, trackable DAG.</li>
                <li><strong>Automation:</strong> Triggers retraining pipelines when data drift is detected or new tables are uploaded.</li>
                <li><strong>Serverless Scale:</strong> Vertex AI Pipelines manages compute resources dynamically, spinning down clusters when tasks complete.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 2 of 10",
            title: "Vertex AI Pipelines & Kubeflow SDK",
            body: `Vertex AI Pipelines compiles Python code compiled using the Kubeflow Pipelines (KFP) SDK:
            <ul>
                <li><strong>DSL Compiler:</strong> KFP converts Python code decorated with <code>@dsl.pipeline</code> into a JSON declarative workflow description.</li>
                <li><strong>Vertex AI execution:</strong> Submits the JSON file to Vertex AI, which runs each component inside independent container tasks, writing outputs to a managed GCS path.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 3 of 10",
            title: "TFX (TensorFlow Extended) Components",
            body: `For TensorFlow workflows, TFX provides highly specialized, standardized components:
            <ul>
                <li><strong>ExampleGen:</strong> Ingests raw files and splits them into training/eval sets.</li>
                <li><strong>SchemaGen & StatisticsGen:</strong> Analyzes feature datatypes and computes baseline metrics.</li>
                <li><strong>Transform:</strong> Runs TF.Transform preprocessing graph generation.</li>
                <li><strong>Evaluator:</strong> Validates model metrics against historical baselines before deployment to prevent model degradation.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 4 of 10",
            title: "Artifact Lineage & Metadata Tracking",
            body: `Vertex ML Metadata tracks pipeline executions:
            <ul>
                <li><strong>Artifacts:</strong> Immutable files produced by components (e.g., Model weights, Evaluator reports, Datasets).</li>
                <li><strong>Lineage Graphs:</strong> Tracks what dataset version trained a specific model artifact, which container processed it, and which endpoint host serves it. Crucial for regulatory compliance and audit trails.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 5 of 10",
            title: "Pipeline Triggering & Automation",
            body: `Scheduling and automating pipeline runs:
            <ul>
                <li><strong>Cloud Scheduler:</strong> Trigger pipeline runs on a calendar schedule (e.g., every Sunday at midnight).</li>
                <li><strong>Cloud Functions + Pub/Sub:</strong> Listen for GCS file additions or database events and trigger a pipeline run programmatically.</li>
                <li><strong>Cloud Composer:</strong> Use Apache Airflow to orchestrate complex data ingestion tasks before executing the training pipeline.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 6 of 10",
            title: "Kubeflow Component Inputs & Outputs",
            body: `Passing data between component nodes:
            <ul>
                <li><strong>Value Parameters:</strong> Small datatypes (strings, integers, floats) passed directly to containers.</li>
                <li><strong>Artifact Paths:</strong> Large objects (datasets, model models) are stored in GCS. Components receive a URI reference and write results back to a GCS path.</li>
                <li><strong>DSL Mapping:</strong> <code>component_b(input_data=component_a.outputs['output_artifact'])</code>.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 7 of 10",
            title: "Conditional Pipeline Branches",
            body: `Making decisions inside pipeline DAG runs:
            <ul>
                <li><strong>dsl.Condition:</strong> Allows executing branches only when specific criteria are met.</li>
                <li><strong>Evaluation gate:</strong> Use a component to compute validation accuracy. If accuracy exceeds the baseline model, execute the deploy branch. If not, terminate without deployment.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 8 of 10",
            title: "TFX Schema Validation",
            body: `Using SchemaGen and Resolver components to maintain data quality:
            <ul>
                <li><strong>Baseline Schema:</strong> Generated during the initial run, defining column types, categorical domains, and value ranges.</li>
                <li><strong>Validation:</strong> The <code>SchemaGen</code> component compares new dataset batches to the baseline. Any datatype shift or missing column triggers an alert, blocking training.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 9 of 10",
            title: "Model Registry Integration in Pipelines",
            body: `Uploading models automatically to the registry:
            <ul>
                <li><strong>upload_model:</strong> A Vertex AI SDK component that registers trained artifacts directly.</li>
                <li><strong>Versioning:</strong> Creates a new version label (e.g., v3) under an existing model name rather than creating a new model resource.</li>
                <li><strong>Metadata:</strong> Associates the pipeline run URI with the model entry for traceability.</li>
            </ul>`
        },
        {
            subtitle: "4. ML Orchestration & Pipelines • Slide 10 of 10",
            title: "Serverless Pipelines vs. Kubernetes",
            body: `Choosing the right orchestration platform:
            <ul>
                <li><strong>Vertex AI Pipelines:</strong> Managed, serverless, zero operations overhead. Best for most teams. Pay-per-run pricing.</li>
                <li><strong>Kubeflow on GKE:</strong> Run pipelines on your own Google Kubernetes Engine (GKE) cluster. Best for teams requiring custom resource scheduling or hybrid environments.</li>
            </ul>`
        }
    ],
    m5_production: [
        {
            subtitle: "5. Production ML & Serving • Slide 1 of 10",
            title: "Online Serving vs. Batch Serving",
            body: `Serving architectures depend on model prediction latency constraints:
            <ul>
                <li><strong>Online Serving (Endpoints):</strong> Low-latency serving. Client sends REST/gRPC payload, model returns predictions in milliseconds. Requires active compute hosts (e.g., Vertex AI Endpoints).</li>
                <li><strong>Batch serving:</strong> High-throughput, offline predictions. Processes millions of rows stored in BigQuery or GCS. Spins down compute instances when prediction completes, saving cost.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 2 of 10",
            title: "Custom Serving Container Configurations",
            body: `For custom inference operations, compile a custom Docker container for Vertex AI:
            <ul>
                <li><strong>Port Rule:</strong> The container must listen on port <code>8080</code> for incoming HTTP requests.</li>
                <li><strong>Routes:</strong> Must implement <code>/health</code> for active health checks, and <code>/predict</code> to receive and return feature payloads.</li>
                <li><strong>Environment:</strong> Pre-load model weights inside the container or download them dynamically from GCS using <code>AIP_STORAGE_URI</code>.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 3 of 10",
            title: "Canary Deployments & Traffic Splitting",
            body: `Minimizing production deployment risks:
            <ul>
                <li><strong>Model Registry:</strong> Vertex AI Model Registry tracks versions (e.g. v1, v2) of your model models.</li>
                <li><strong>Traffic Splitting:</strong> When deploying a new model version, deploy it to the same Endpoint instance and allocate a minor traffic share (e.g., 10%).</li>
                <li><strong>Canary validation:</strong> Monitor v2 prediction metrics. If metrics remain healthy, ramp traffic allocations to 100% and undeploy the older model.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 4 of 10",
            title: "Vertex AI Model Monitoring",
            body: `Tracks models in production to detect performance drops:
            <ul>
                <li><strong>Training-Serving Skew:</strong> Production feature inputs differ from training data distributions. Measured using baseline TFRecords.</li>
                <li><strong>Prediction Drift:</strong> Feature inputs shift day-over-day in production (data drift).</li>
                <li><strong>PSI Metric:</strong> Population Stability Index (PSI) measures distribution differences. A threshold of <strong>PSI >= 0.20</strong> triggers system alerts and automated pipeline runs.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 5 of 10",
            title: "Autoscaling Serving Endpoints",
            body: `Configuring scaling properties on Vertex AI Endpoints:
            <ul>
                <li><strong>Min/Max Nodes:</strong> Setting minimum nodes to 0 allows scaling down to zero when idle (saving cost), but introduces cold start latencies.</li>
                <li><strong>Target Metrics:</strong> Autoscales nodes dynamically based on average CPU usage or GPU utilization.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 6 of 10",
            title: "Model Explanations (Feature Attribution)",
            body: `Understand why a model made a specific prediction:
            <ul>
                <li><strong>Feature Attributions:</strong> Returns a score showing how much each input feature contributed to the prediction value.</li>
                <li><strong>Methods:</strong> Supports <strong>Integrated Gradients</strong> (best for deep learning neural networks) and <strong>Sampled Shapley</strong> (best for ensemble tree models).</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 7 of 10",
            title: "Vertex AI Model Registry Governance",
            body: `Governance patterns for model deployment:
            <ul>
                <li><strong>Aliases:</strong> Set mutable labels like <code>@default</code> or <code>@production</code> to route endpoint calls to active versions without changing client code.</li>
                <li><strong>Sharing:</strong> Share models across multiple GCP projects to decouple model training from web operations.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 8 of 10",
            title: "GPU Serving Optimization",
            body: `Maximize inference throughput:
            <ul>
                <li><strong>TensorRT Compilation:</strong> Optimizes neural networks for GPU deployment, converting weights to FP16 or INT8 precision.</li>
                <li><strong>Triton Inference Server:</strong> Serve PyTorch and TensorFlow models concurrently with dynamic batching, maximizing GPU hardware utilization.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 9 of 10",
            title: "Model Monitoring Alert Thresholds",
            body: `Configuring alerts for production drift:
            <ul>
                <li><strong>Skew Thresholds:</strong> Set alerting triggers on specific features (e.g. <code>L1_distance > 0.1</code>).</li>
                <li><strong>Alert Actions:</strong> Monitors publish warning messages to Pub/Sub, triggering automated functions to initiate retraining runs.</li>
            </ul>`
        },
        {
            subtitle: "5. Production ML & Serving • Slide 10 of 10",
            title: "Edge Deployment (TF Lite & Edge TPU)",
            body: `Deploying models to IoT devices and localized hardware:
            <ul>
                <li><strong>TensorFlow Lite:</strong> Converts SavedModel files into flat buffers, reducing size and memory foot-print.</li>
                <li><strong>Quantization:</strong> Compresses model weights from float32 to int8 to enable execution on Edge TPUs (like Coral dev boards).</li>
            </ul>`
        }
    ],
    m6_genai: [
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 1 of 10",
            title: "Foundation Models & Model Garden",
            body: `Generative AI on Google Cloud is structured around Foundation Models:
            <ul>
                <li><strong>Vertex AI Model Garden:</strong> A catalog for discovering and deploying Google first-party models (Gemini), open-source models (Gemma, Llama), and third-party models.</li>
                <li><strong>Vertex AI Studio:</strong> Interactive web playgrounds for prototyping prompt configurations (Freeform, Chat, and Structured Prompts) before API integration.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 2 of 10",
            title: "Transformer Attentions & Parameters",
            body: `How LLMs generate text token sequences:
            <ul>
                <li><strong>Self-Attention:</strong> Math weights representing how words relate in context.</li>
                <li><strong>Temperature (Temp):</strong> Scales next-token probability distribution. High Temperature (1.0+) increases creativity and variety. Temperature close to 0 enforces deterministic output.</li>
                <li><strong>Top-K & Top-P:</strong> Truncate vocabulary to highest-probability tokens. Top-P dynamically scales vocabulary boundary until cumulative sum exceeds P.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 3 of 10",
            title: "Retrieval-Augmented Generation (RAG)",
            body: `To bypass LLM knowledge limits and hallucinations without updating model weights:
            <ul>
                <li><strong>Vector Search:</strong> Convert user queries into numerical embeddings and search high-dimensional indexes (like Vertex AI Vector Search).</li>
                <li><strong>Context Injection:</strong> Fetch relevant enterprise documents and append them directly to the prompt as context, grounding the LLM output.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 4 of 10",
            title: "Parameter-Efficient Fine-Tuning (PEFT)",
            body: `When prompt sizing is too small to convey style or domain guidelines:
            <ul>
                <li><strong>PEFT/LoRA:</strong> Freezes the baseline billions of parameters of the model and introduces low-rank matrices to train adapter weights.</li>
                <li><strong>Advantages:</strong> Significantly reduces GPU memory requirement and training time during Vertex AI custom tuning pipeline jobs.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 5 of 10",
            title: "Prompt Engineering Patterns",
            body: `Structuring inputs to guide model outputs:
            <ul>
                <li><strong>Few-Shot:</strong> Provide example formats to show style (e.g. input/output pairs).</li>
                <li><strong>Chain-of-Thought:</strong> Force the model to output reasoning steps (e.g., 'think step-by-step') to improve logic performance.</li>
                <li><strong>ReAct Pattern:</strong> Interleaves reasoning with external actions (like calling database APIs) to solve complex user requests.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 6 of 10",
            title: "Vertex AI Studio Playgrounds",
            body: `Prototyping tools in the Vertex AI console:
            <ul>
                <li><strong>Freeform Prompts:</strong> General workspace for prompt design.</li>
                <li><strong>Chat Prompts:</strong> Simulates messaging histories, allowing system instruction definitions.</li>
                <li><strong>Structured Prompts:</strong> Formats examples in tabular rows, perfect for classification or structured output generations.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 7 of 10",
            title: "GenAI SDK Integration",
            body: `Interacting with Gemini programmatically:
            <ul>
                <li><strong>Python SDK:</strong> Use the <code>google-genai</code> library to initialize model clients.</li>
                <li><strong>Schema Enforcement:</strong> Pass JSON schemas to force model outputs to strictly comply with target API structures.</li>
                <li><strong>System Instructions:</strong> Define global instructions that govern model tone and behavior throughout sessions.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 8 of 10",
            title: "Model Distillation on Vertex AI",
            body: `Compressing Large Language Models:
            <ul>
                <li><strong>Teacher-Student Setup:</strong> Uses a larger model (e.g., Gemini 1.5 Pro) to generate training targets.</li>
                <li><strong>Distillation Job:</strong> Trains a smaller student model (e.g., Gemma 2B) to mirror the larger model's behaviors, reducing inference compute costs.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 9 of 10",
            title: "GenAI Security & Safety Filters",
            body: `Configuring safety boundaries on Vertex AI:
            <ul>
                <li><strong>Safety Categories:</strong> Adjust block thresholds for Hate Speech, Harassment, Sexually Explicit Content, and Dangerous Content.</li>
                <li><strong>Threshold levels:</strong> Set levels to block low, medium, or high probability violations, returning empty/blocked payload statuses when violated.</li>
            </ul>`
        },
        {
            subtitle: "6. GenAI & LLMs on GCP • Slide 10 of 10",
            title: "Vertex AI Agent Builder",
            body: `Build Conversational Enterprise Agents:
            <ul>
                <li><strong>Enterprise Search:</strong> Connects models to Cloud Storage or BigQuery folders to run grounded internal searches.</li>
                <li><strong>Dialog Flow integration:</strong> Builds conversational logic scripts connected to company database APIs to run customer workflows (such as reservation updates).</li>
            </ul>`
        }
    ]
};

// --- Practice Exam Question Bank (25 Scenario-Based Questions) ---
const quizQuestions = [
    {
        question: "A company needs to transform raw text files into TFRecord training datasets for a custom TensorFlow deep learning model. The transformation pipeline must handle massive batch processing daily and scale up to 100 compute nodes. Which GCP tool is recommended?",
        choices: [
            "BigQuery ML logistic regression pipelines.",
            "Cloud Dataflow running an Apache Beam pipeline to process text inputs and export TFRecord files to Cloud Storage.",
            "Cloud Dataproc running a legacy MapReduce job using HDFS cluster storage.",
            "Vertex AI Studio prompting workflows."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Cloud Dataflow is Google's managed Apache Beam processing engine, designed for high-scale batch and streaming data preprocessing. Running Beam transformations and exporting to TFRecords (TensorFlow's native dataset format) on GCS is the standard GCP pipeline pattern."
    },
    {
        question: "You are designing a neural network model to predict product churn. Your training dataset contains features with widely varying scales. To ensure rapid optimizer convergence, you need to standardize your tabular numerical features. How should you design the preprocessing steps to avoid train-serve skew?",
        choices: [
            "Use Cloud Dataflow to calculate mean and variance, transform the training data, and copy the Python preprocessing functions into the serving API.",
            "Use TF.Transform within a TFX pipeline to calculate dataset constants and generate a preprocessing graph that is exported directly as part of the model artifact.",
            "Run SQL operations in BigQuery during prediction fetches to scale feature inputs.",
            "Set the model's prediction serving Temperature to 0.0 to stabilize output scaling."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Preprocessing steps must be identical during training and serving to prevent skew. TF.Transform generates a preprocessing graph that executes as a model layer, ensuring both training records and raw serving requests are scaled with the exact same mean/variance constants."
    },
    {
        question: "Your team is building a deep learning transformer model using PyTorch. The model weights require 90GB of memory, which exceeds the memory capacity of your largest single GPU worker node. How should you orchestrate the custom training job on Vertex AI?",
        choices: [
            "Use AutoML Tabular to automate model generation.",
            "Orchestrate a distributed training job using Model Parallelism to split the model layers across multiple GPU worker nodes.",
            "Configure Data Parallelism to replicate model weights across multiple workers.",
            "Deploy a Vertex AI Reduction Server to scale down the PyTorch model weights to 16GB."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. When the model itself exceeds single-node memory, you must use Model Parallelism to distribute the network structure across multiple nodes. Data Parallelism (C) distributes data batches but replicates the full model on each worker, which would crash due to Out-Of-Memory (OOM) errors. Reduction Server (D) aggregates gradients, it does not compress model parameters."
    },
    {
        question: "You want to run a custom training pipeline on Vertex AI to optimize a neural network. You want Vizier to automatically tune the learning rate log scale and dropout rate hyperparameters to minimize validation loss. What is the recommended configuration?",
        choices: [
            "Run a custom container that executes Vizier local grid-search scripts inside the container.",
            "Configure a Vertex AI Hyperparameter Tuning Job specifying the parameter types, boundaries (log/linear), target metric, and optimization algorithm (like Bayesian Optimization).",
            "Use BigQuery ML hyperparameter tuning loops inside SQL procedures.",
            "Deploy the model to an online endpoint and adjust parameters manually using canary traffic splits."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Vertex AI Hyperparameter Tuning Jobs run as a managed service, using Vizier to orchestrate trials. By defining parameter bounds, scales, and target metric thresholds, Vertex AI spins up training tasks and optimizes hyperparameter search space using Bayesian Optimization."
    },
    {
        question: "You are setting up an MLOps pipeline on Google Cloud. You must ensure that every trained model can be audited to trace the exact dataset version used during training. How should you implement this?",
        choices: [
            "Write model metadata details directly into logs and query Cloud Logging.",
            "Use Vertex ML Metadata inside your Vertex AI Pipelines to track pipeline execution runs and establish an artifact lineage graph.",
            "Use a custom naming convention for output model files stored in Cloud Storage buckets.",
            "Store all models in a local SQLite database on a Compute Engine instance."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Vertex ML Metadata tracks artifact lineage. When running Vertex AI Pipelines (using Kubeflow), metadata links datasets, intermediate outputs, and the final model. This creates a searchable lineage graph that traces model provenance for compliance."
    },
    {
        question: "You are deploying a custom TensorFlow model for online serving. You build a custom container for Vertex AI Prediction. Which configuration rules must your custom container satisfy?",
        choices: [
            "Listen on port 80, accept POST requests at /inference, and run on a GPU node.",
            "Listen on port 8080, implement /health and /predict endpoints, and return JSON prediction arrays.",
            "Run on port 8888, load model weights from a local volume, and support gRPC calls.",
            "Store model weights in an internal SQLite database and expose port 443."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Vertex AI Prediction custom serving containers must listen on port 8080 for incoming HTTP traffic, respond to GET requests on /health with HTTP 200, and handle POST requests containing inference inputs on /predict."
    },
    {
        question: "You have deployed a new recommendation model (v2) to a Vertex AI Endpoint. You want to validate its real-world performance against the active model (v1) before routing all customer requests to it. What MLOps serving pattern should you use?",
        choices: [
            "Undeploy v1 and monitor customer feedback channels.",
            "Deploy v2 to the same Endpoint instance and configure a Traffic Split to route 10% of traffic to v2 and 90% to v1, monitoring metrics.",
            "Build a separate endpoint, deploy v2, and write client code to randomly route requests.",
            "Run a batch prediction job on the serving endpoints daily."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Vertex AI Endpoints allow multiple models to be deployed to a single endpoint. You can configure a traffic split (canary deployment) to direct a small fraction of traffic to the new version, validating performance metrics before full migration."
    },
    {
        question: "You need to monitor a production prediction model for feature distribution shifts over time. Which Vertex AI service should you implement, and what statistical metrics does it track?",
        choices: [
            "Vertex AI Pipelines tracking ROC-AUC metrics.",
            "Vertex AI Model Monitoring to detect prediction drift and training-serving skew, tracking the Population Stability Index (PSI) metric.",
            "Vizier hyperparameter search tracking trial evaluation metrics.",
            "Artifact Registry tracking container dependencies."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Vertex AI Model Monitoring is Google's managed solution to track feature drift and training-serving skew. It uses the Population Stability Index (PSI) to measure the divergence between production prediction requests and baseline training distributions, alerting engineers if PSI exceeds a set threshold."
    },
    {
        question: "A company wants to perform sentiment analysis on customer support emails. They want a fast deployment, do not have a dedicated data science team, and do not need to customize model weights. What is the most appropriate GCP choice?",
        choices: [
            "Train a custom BERT model using PyTorch on a GPU node cluster.",
            "Call the pre-trained Google Cloud Natural Language API directly.",
            "Write a custom TF.Transform script inside a Dataflow job.",
            "Build a Kubeflow pipeline to train a Scikit-Learn classifier."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. The Natural Language API is a pre-trained out-of-the-box solution that handles sentiment analysis, entity extraction, and syntax analysis with a simple REST request. Since they have no data science team and don't need custom weights, this is the most cost-effective and rapid choice."
    },
    {
        question: "You are modeling structured customer transaction tables in BigQuery. You want to quickly prototype a binary classifier to predict churn using SQL queries. Which tool is recommended?",
        choices: [
            "AutoML Vision API.",
            "BigQuery ML (BQML) using logistic regression.",
            "Vertex AI Pipelines with TensorFlow components.",
            "Vizier hyperparameter optimization."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. BigQuery ML allows developers to train and evaluate ML models directly inside BigQuery using standard SQL syntax. For rapid prototyping on tabular data, BQML is the ideal tool as it avoids data export latency."
    },
    {
        question: "You are training a large NLP model on a Cloud TPU. You notice that the training speed drops significantly and the logs indicate high data serialization overhead. What adjustment should you make to optimize TPU training performance?",
        choices: [
            "Use PyTorch dynamic tensors with varying batch dimensions.",
            "Use TFRecord file structures on GCS and configure static tensor input dimensions to maximize XLA compiler optimization.",
            "Increase the serving Temperature parameter to 1.5.",
            "Deploy a custom container to route predictions to port 8080."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. TPUs utilize the XLA compiler to optimize matrix math. XLA requires static tensor dimensions; dynamic shapes force model recompilations, hurting performance. Storing training data in TFRecords on GCS ensures fast ingestion streaming, eliminating serialization bottlenecks."
    },
    {
        question: "Your company has deployed a model to serve real-time predictions. The prediction features include user geolocation coordinates. Over the holidays, user check-in patterns shift, causing a shift in feature coordinates. What type of model degradation is this, and how can Vertex AI detect it?",
        choices: [
            "Concept drift; detected by evaluating ROC-AUC scores on custom metrics.",
            "Prediction/Data drift; detected by Vertex AI Model Monitoring calculating PSI changes over production request windows.",
            "Train-serve skew; detected by comparing preprocessing graphs in TF.Transform.",
            "Vizier optimization drift; detected by running hyperparameter trials."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Prediction drift (or data drift) occurs when the statistical distribution of feature inputs changes over time in production. Vertex AI Model Monitoring calculates the Population Stability Index (PSI) daily to detect these variations and alert engineers."
    },
    {
        question: "You need to frame an ML solution to predict whether a customer will click an ad. What type of ML task is this, and what metrics should you use to evaluate performance if ad clicks are highly rare?",
        choices: [
            "Unsupervised clustering; evaluate using silhouette coefficients.",
            "Supervised binary classification; evaluate using Precision-Recall AUC (PR-AUC) or F1-Score.",
            "Regression; evaluate using Mean Squared Error (MSE).",
            "Generative AI; evaluate using Temperature parameters."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Predicting a click (Yes/No) is a binary classification task. In highly imbalanced datasets (clicks are rare), standard accuracy can be misleadingly high (e.g. 99% accuracy by predicting 'No' to everything). PR-AUC or F1-Score are appropriate metrics for class imbalance."
    },
    {
        question: "You are building a custom preprocessing pipeline using Cloud Dataflow. You want to extract features from images stored in GCS and write the TFRecords. During the Beam job, you encounter out-of-memory issues on the Dataflow workers. How should you adjust your pipeline configuration?",
        choices: [
            "Change Dataflow worker types to machine profiles with larger memory capacities, and adjust worker auto-scaling properties.",
            "Replace Dataflow with BigQuery ML queries.",
            "Deploy the model to an online endpoint and run custom preprocessing during predictions.",
            "Run Vizier hyperparameter tuning to optimize batch sizes."
        ],
        correctAnswer: 0,
        explanation: "Correct answer: Option A. Processing large image files in memory during Beam transformations can trigger OOM errors. Scaling the memory capacity of the Dataflow worker instances (e.g., using n1-highmem machines) resolves this resource bottleneck."
    },
    {
        question: "How does training-serving skew differ from prediction drift in MLOps monitoring?",
        choices: [
            "Skew is measured in training code syntax, while drift is measured in serving code syntax.",
            "Skew is a statistical difference between training data and serving data distributions, while drift is a statistical shift in serving data over time.",
            "Skew is managed by Vizier, while drift is managed by Kubeflow pipelines.",
            "Skew affects LLM output parameters, while drift affects model endpoint ports."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Training-serving skew is the distribution difference between training inputs and serving inputs. Prediction drift is the shift of serving data distributions over time (e.g., season-to-season customer behavior shifts)."
    },
    {
        question: "Which TFX component is responsible for evaluating model quality metrics (like validation loss or ROC-AUC) against historical production models before deploying to the registry?",
        choices: [
            "ExampleGen",
            "Evaluator",
            "Transform",
            "SchemaGen"
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. The TFX Evaluator component analyzes model performance on evaluation sets and compares the metrics against baseline models (like the active production model). If the new model performs better and meets safety thresholds, it is approved for registry deployment."
    },
    {
        question: "You need to write a Kubeflow pipeline using the KFP SDK to automate training. Which decorator is used to compile a Python function into an independent pipeline task component?",
        choices: [
            "@dsl.pipeline",
            "@dsl.component",
            "@aiplatform.training_job",
            "@custom_container"
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. The `@dsl.component` decorator compiles a python function into a Kubeflow component container task. The `@dsl.pipeline` decorator (A) is used to define the overall DAG structure linking those components."
    },
    {
        question: "A company wants to generate creative marketing copy using a Large Language Model. They want to maximize the variety and uniqueness of the generated text, allowing for diverse styles. Which hyperparameter adjustments should you make?",
        choices: [
            "Set Temperature to 0.0 and Top-K to 1.",
            "Increase Temperature (e.g. 1.2) and set Top-P to 0.95.",
            "Set Temperature to 0.01 and Top-P to 0.1.",
            "Disable Top-K/Top-P and lower Temperature to 0."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Higher Temperatures flatten the probability distribution of vocabulary tokens, allowing less common tokens to be selected. High Top-P (0.95) keeps a wide vocabulary of active tokens, maximizing variation."
    },
    {
        question: "You are building a customer support bot using a foundation LLM. You want to ground the LLM's answers using internal documentation files that are updated continuously. How should you design the system?",
        choices: [
            "Run a daily Full Fine-Tuning pipeline on the LLM using hospital logs.",
            "Implement Retrieval-Augmented Generation (RAG) by converting documents to vector embeddings, searching them at runtime using Vertex AI Vector Search, and injecting context into the prompt.",
            "Set Temperature to 2.0 to make the model explore its internal knowledge.",
            "Train a custom regression model on the documentation."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. RAG is the standard design to inject external context into prompts at runtime, ensuring grounded, accurate, and up-to-date answers without retraining models."
    },
    {
        question: "When tuning an LLM, you decide to freeze the base model parameters and train a small set of added adaptor weights. What is this tuning methodology called?",
        choices: [
            "Full Fine-Tuning",
            "Parameter-Efficient Fine-Tuning (PEFT/LoRA)",
            "Retrieval-Augmented Generation (RAG)",
            "Vizier Hyperparameter Search"
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. PEFT (including LoRA - Low-Rank Adaptation) freezes the primary weights of a foundation model and inserts small trainable adapter matrices, reducing GPU memory and training time."
    },
    {
        question: "Your ML model needs to serve batch predictions on a 10TB dataset in BigQuery every Sunday night. The predictions must be exported to a Cloud Storage bucket. Latency is not a constraint, but cost efficiency is. How should you serve predictions?",
        choices: [
            "Deploy the model to a Vertex AI Endpoint with 10 active GPU nodes.",
            "Create a Vertex AI Batch Prediction Job, specifying BigQuery as the source and GCS as the target.",
            "Host the model in a custom container on Compute Engine and write client loops.",
            "Use BigQuery ML online API calls."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Batch Prediction Jobs are designed for high-throughput offline predictions. They spin up resources to run predictions and automatically terminate them when complete, minimizing cost compared to continuous online endpoint hosting."
    },
    {
        question: "You want to design a pipeline to update model features in real-time. Features must be available for online model calls with a latency of less than 10ms. Which GCP components should you use?",
        choices: [
            "BigQuery -> BigQuery ML -> GCS",
            "Pub/Sub -> Cloud Dataflow -> Vertex AI Feature Store Online serving API",
            "Cloud Storage -> Dataproc -> Cloud SQL",
            "Cloud Composer -> Cloud Functions -> BigQuery ML"
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Pub/Sub handles incoming real-time telemetry events, Dataflow cleanses and transforms the data stream, and feeds features into Vertex AI Feature Store's online store serving (Bigtable), which serves low-latency prediction fetches."
    },
    {
        question: "During training of a computer vision model, you notice that validation loss is increasing while training loss decreases. How should you address this issue?",
        choices: [
            "The model is underfitting; increase model capacity and training epochs.",
            "The model is overfitting; add regularization (L1/L2, dropout) or gather more training data.",
            "The learning rate is too low; run a Vizier study to decrease it.",
            "Set the generation Temperature parameter to 0.0."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Divergence where training loss decreases but validation loss increases is the signature of overfitting (model memorizing training data). Remediations include regularization, early stopping, or increasing dataset size."
    },
    {
        question: "You are setting up Vertex AI Model Monitoring. You want to detect if production prediction inputs deviate from the baseline dataset used to train the model. What is this configuration called, and what baseline input file format is expected?",
        choices: [
            "Prediction drift; baseline input is daily prediction logs.",
            "Training-Serving Skew; baseline input is the training dataset (e.g. TFRecords stored in GCS).",
            "Vizier trial; baseline input is the optimal hyperparameter loss value.",
            "Pipeline validation; baseline input is the Kubeflow JSON."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. Testing production inputs against original training data is called Training-Serving Skew monitoring. It requires providing the original training data path (like GCS TFRecords) as a baseline comparison source."
    },
    {
        question: "What does a Population Stability Index (PSI) value of 0.25 indicate in a Vertex AI Model Monitoring report?",
        choices: [
            "Excellent model convergence with minimal loss.",
            "Significant distribution shift (drift/skew), indicating that retraining is highly recommended to prevent performance drops.",
            "The model is 25% completed with its training cycles.",
            "Traffic splits should be adjusted to route 25% of requests to the canary model."
        ],
        correctAnswer: 1,
        explanation: "Correct answer: Option B. PSI measures distribution changes. A value < 0.1 indicates no shift; 0.1 to 0.2 indicates moderate shift; and PSI >= 0.2 indicates significant shift, meaning the input features have drifted and the model should be retrained."
    }
];

// --- Vocabulary Database for the LLM Sandbox Simulator ---
const promptVocabulary = {
    "Machine learning is": [
        { token: "powerful", baseProb: 0.40 },
        { token: "fun", baseProb: 0.25 },
        { token: "difficult", baseProb: 0.15 },
        { token: "changing", baseProb: 0.10 },
        { token: "revolutionary", baseProb: 0.05 },
        { token: "overrated", baseProb: 0.03 },
        { token: "everywhere", baseProb: 0.02 }
    ],
    "Generative AI will": [
        { token: "transform", baseProb: 0.45 },
        { token: "automate", baseProb: 0.20 },
        { token: "replace", baseProb: 0.15 },
        { token: "improve", baseProb: 0.10 },
        { token: "disrupt", baseProb: 0.06 },
        { token: "fail", baseProb: 0.03 },
        { token: "evolve", baseProb: 0.01 }
    ],
    "Vertex AI Feature Store solves": [
        { token: "skew", baseProb: 0.40 },
        { token: "latency", baseProb: 0.30 },
        { token: "leakage", baseProb: 0.18 },
        { token: "redundancy", baseProb: 0.08 },
        { token: "bugs", baseProb: 0.03 },
        { token: "cost", baseProb: 0.01 }
    ],
    "The transformer attention mechanism": [
        { token: "relates", baseProb: 0.35 },
        { token: "calculates", baseProb: 0.25 },
        { token: "parallelizes", baseProb: 0.20 },
        { token: "improves", baseProb: 0.12 },
        { token: "replaces", baseProb: 0.06 },
        { token: "fails", baseProb: 0.02 }
    ]
};

// --- App Initialization & Navigation ---
document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    initSlides();
    initFeatureStoreSimulator();
    initVizierSimulator();
    initModelMonitorSimulator();
    initLLMSandbox();
    initQuiz();
    updateGlobalProgress();
});

function initNavigation() {
    const navLinks = document.querySelectorAll('.nav-link');
    const sections = document.querySelectorAll('.content-section');
    
    navLinks.forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const targetSection = link.getAttribute('data-target');
            
            navLinks.forEach(l => l.classList.remove('active'));
            sections.forEach(s => s.classList.remove('active'));
            
            link.classList.add('active');
            const targetEl = document.getElementById(targetSection);
            if (targetEl) {
                targetEl.classList.add('active');
            }
            
            state.currentSection = targetSection;
            
            if (targetSection === 'practice-exam' && !state.quiz.answered && state.quiz.timerInterval === null) {
                startQuizTimer();
            }
        });
    });
}

function updateGlobalProgress() {
    const modules = ['m1_fundamentals', 'm2_features', 'm3_development', 'm4_pipelines', 'm5_production', 'm6_genai'];
    let modulesProgressSum = 0;
    
    modules.forEach((mod, idx) => {
        const modKey = `m${idx + 1}`;
        const activeSlideIdx = state.slides[mod].current;
        const totalSlides = state.slides[mod].total;
        
        const progressVal = ((activeSlideIdx + 1) / totalSlides) * 100;
        state.progress[modKey] = Math.max(state.progress[modKey], progressVal);
        modulesProgressSum += state.progress[modKey];
    });
    
    const quizAnsweredCount = state.quiz.userAnswers.length;
    const quizProgressVal = (quizAnsweredCount / quizQuestions.length) * 100;
    state.progress.quiz = Math.max(state.progress.quiz, quizProgressVal);
    
    const totalProgress = (modulesProgressSum + state.progress.quiz) / 7;
    
    const fillEl = document.getElementById('global-progress-fill');
    const textEl = document.getElementById('global-progress-text');
    if (fillEl) fillEl.style.width = `${totalProgress}%`;
    if (textEl) textEl.textContent = `${Math.round(totalProgress)}%`;
    
    updateDashboardProgressLabels();
}

function updateDashboardProgressLabels() {
    for (let i = 1; i <= 6; i++) {
        const label = document.getElementById(`dash-progress-m${i}`);
        if (label) {
            label.textContent = `${Math.round(state.progress[`m${i}`])}%`;
        }
    }
    const examLabel = document.getElementById('dash-progress-exam');
    if (examLabel) examLabel.textContent = `${Math.round(state.progress.quiz)}%`;
}

// --- Presentation Deck Logic ---
function initSlides() {
    const modules = ['m1_fundamentals', 'm2_features', 'm3_development', 'm4_pipelines', 'm5_production', 'm6_genai'];
    
    modules.forEach(mod => {
        renderSlide(mod);
        
        const prevBtn = document.getElementById(`btn-${mod}-prev`);
        const nextBtn = document.getElementById(`btn-${mod}-next`);
        
        if (prevBtn) {
            prevBtn.addEventListener('click', () => {
                if (state.slides[mod].current > 0) {
                    state.slides[mod].current--;
                    renderSlide(mod);
                    updateGlobalProgress();
                }
            });
        }
        
        if (nextBtn) {
            nextBtn.addEventListener('click', () => {
                if (state.slides[mod].current < state.slides[mod].total - 1) {
                    state.slides[mod].current++;
                    renderSlide(mod);
                    updateGlobalProgress();
                }
            });
        }
    });
}

function renderSlide(mod) {
    const currentIdx = state.slides[mod].current;
    const slides = slideData[mod];
    const slide = slides[currentIdx];
    
    const container = document.getElementById(`slide-deck-${mod}`);
    if (!container) return;
    
    container.innerHTML = `
        <div class="slide active">
            <span class="slide-subtitle">${slide.subtitle}</span>
            <h2 class="slide-title">${slide.title}</h2>
            <div class="slide-body">${slide.body}</div>
        </div>
    `;
    
    const counter = document.getElementById(`counter-${mod}`);
    if (counter) {
        counter.textContent = `Slide ${currentIdx + 1} of ${state.slides[mod].total}`;
    }
    
    const prevBtn = document.getElementById(`btn-${mod}-prev`);
    const nextBtn = document.getElementById(`btn-${mod}-next`);
    
    if (prevBtn) prevBtn.disabled = currentIdx === 0;
    if (nextBtn) nextBtn.disabled = currentIdx === state.slides[mod].total - 1;
}

// --- Vertex AI Feature Store Simulator ---
function initFeatureStoreSimulator() {
    const triggerBtn = document.getElementById('fs-simulate-btn');
    if (!triggerBtn) return;
    
    triggerBtn.addEventListener('click', () => {
        const ingestionType = document.getElementById('fs-ingest-select').value;
        const servingTarget = document.getElementById('fs-serve-select').value;
        runFeatureStoreSimulation(ingestionType, servingTarget);
    });
}

function runFeatureStoreSimulation(ingestion, serving) {
    const logConsole = document.getElementById('fs-log-console');
    const latencyVal = document.getElementById('fs-latency-val');
    
    document.querySelectorAll('.fs-node, .fs-node-serving').forEach(el => {
        el.classList.remove('active', 'active-online', 'active-offline');
    });
    document.querySelectorAll('.fs-line').forEach(el => {
        el.classList.remove('active-online', 'active-offline');
    });
    
    logConsole.innerHTML = `<span class="code-comment">// Connecting to Vertex AI Feature Store client...</span>\n`;
    
    setTimeout(() => {
        let nodeIngestId = ingestion === 'batch' ? 'node-bq' : 'node-pubsub';
        let lineIngestId = ingestion === 'batch' ? 'line-bq-fs' : 'line-pubsub-fs';
        
        const nodeIngest = document.getElementById(nodeIngestId);
        const nodeFs = document.getElementById('node-fs');
        const lineIngest = document.getElementById(lineIngestId);
        
        if (nodeIngest) nodeIngest.classList.add('active');
        if (nodeFs) nodeFs.classList.add('active');
        
        let classColor = serving === 'online' ? 'active-online' : 'active-offline';
        if (lineIngest) lineIngest.classList.add(classColor);
        
        if (ingestion === 'batch') {
            logConsole.innerHTML += `<span style="color: #fda085;">[INFO] Ingesting features from BigQuery tables. Triggering Point-in-Time view sync.</span>\n`;
            logConsole.innerHTML += `gcloud vertex-ai feature-groups create ... --bq-source=bq://my-proj.ds.table\n`;
        } else {
            logConsole.innerHTML += `<span style="color: #43e97b;">[INFO] Streaming stream detected. Syncing Pub/Sub telemetry features via Dataflow pipeline.</span>\n`;
            logConsole.innerHTML += `PythonSDK: FeatureStore.write_feature_values(entity_id="user_91", values={"click_rate": 0.82})\n`;
        }
        logConsole.scrollTop = logConsole.scrollHeight;
    }, 500);
    
    setTimeout(() => {
        let nodeServingId = serving === 'online' ? 'node-online' : 'node-offline';
        let lineServingId = serving === 'online' ? 'line-fs-online' : 'line-fs-offline';
        
        const nodeServing = document.getElementById(nodeServingId);
        const lineServing = document.getElementById(lineServingId);
        
        let classColor = serving === 'online' ? 'active-online' : 'active-offline';
        
        if (nodeServing) nodeServing.classList.add(classColor);
        if (lineServing) lineServing.classList.add(classColor);
        
        if (serving === 'online') {
            latencyVal.textContent = "6.4 ms";
            latencyVal.style.color = "var(--accent-green)";
            
            logConsole.innerHTML += `<span style="color: #43e97b;">[SUCCESS] Online Serving API response fetched in 6.4ms (backed by Bigtable key-value storage)</span>\n`;
            logConsole.innerHTML += `{\n  "entityId": "user_49810",\n  "features": {\n    "last_active": "2026-06-10T21:30:00Z",\n    "purchase_frequency": 0.14,\n    "churn_risk": 0.02\n  }\n}\n`;
        } else {
            latencyVal.textContent = "2.8 min";
            latencyVal.style.color = "#fda085";
            
            logConsole.innerHTML += `<span style="color: #fda085;">[SUCCESS] Offline Export initiated to BigQuery. Completed point-in-time snapshot join.</span>\n`;
            logConsole.innerHTML += `Exported: gs://my-bucket/training-sets/export-1685.csv\n`;
        }
        logConsole.scrollTop = logConsole.scrollHeight;
    }, 2000);
}

// --- Vertex AI Vizier Hyperparameter Tuning Simulator ---
function initVizierSimulator() {
    const runBtn = document.getElementById('vizier-run-btn');
    if (!runBtn) return;
    
    runBtn.addEventListener('click', () => {
        const alg = document.getElementById('vizier-algorithm').value;
        runVizierStudySimulation(alg);
    });
}

function runVizierStudySimulation(alg) {
    const grid = document.getElementById('vizier-dot-grid');
    const logConsole = document.getElementById('vizier-log');
    const bestLossEl = document.getElementById('vizier-best-loss');
    
    const labels = grid.querySelectorAll('.vizier-grid-label-x, .vizier-grid-label-y');
    grid.innerHTML = '';
    labels.forEach(l => grid.appendChild(l));
    
    bestLossEl.textContent = "--";
    logConsole.innerHTML = `<span class="code-comment">// Creating hyperparameter tuning study in Vizier...</span>\n`;
    
    let bestLoss = 999.0;
    let trialCount = 10;
    let currentTrial = 1;
    
    const optX = 55;
    const optY = 40;
    
    function runNextTrial() {
        if (currentTrial > trialCount) {
            logConsole.innerHTML += `<span style="color: var(--accent-green);">[STUDY COMPLETE] Optimal trial found! Minimized Eval Loss: ${bestLoss.toFixed(4)}</span>\n`;
            logConsole.scrollTop = logConsole.scrollHeight;
            return;
        }
        
        let lr, bs, x, y, loss;
        
        if (alg === 'bayesian') {
            let convergenceFactor = (trialCount - currentTrial) / trialCount;
            let noiseX = (Math.random() - 0.5) * 60 * convergenceFactor;
            let noiseY = (Math.random() - 0.5) * 60 * convergenceFactor;
            x = Math.max(10, Math.min(90, optX + noiseX));
            y = Math.max(10, Math.min(90, optY + noiseY));
        } else {
            x = 10 + Math.random() * 80;
            y = 10 + Math.random() * 80;
        }
        
        let dist = Math.sqrt(Math.pow(x - optX, 2) + Math.pow(y - optY, 2));
        loss = 0.05 + (dist / 100) * 0.8 + Math.random() * 0.05;
        
        if (loss < bestLoss) {
            bestLoss = loss;
            bestLossEl.textContent = bestLoss.toFixed(4);
            bestLossEl.style.color = "var(--accent-green)";
        }
        
        lr = Math.pow(10, -5 + (x / 100) * 4);
        bs = Math.round(Math.pow(2, 3 + (y / 100) * 6));
        
        let dotColor = loss < 0.2 ? 'var(--accent-green)' : (loss < 0.4 ? 'var(--accent-blue)' : '#fda085');
        
        const dot = document.createElement('div');
        dot.className = 'vizier-trial-dot';
        dot.style.left = `${x}%`;
        dot.style.bottom = `${y}%`;
        dot.style.backgroundColor = dotColor;
        grid.appendChild(dot);
        
        dot.classList.add('active');
        setTimeout(() => dot.classList.remove('active'), 500);
        
        logConsole.innerHTML += `Trial ${currentTrial}: LearningRate=${lr.toExponential(2)}, BatchSize=${bs} -> EvalLoss=${loss.toFixed(4)}\n`;
        logConsole.scrollTop = logConsole.scrollHeight;
        
        currentTrial++;
        setTimeout(runNextTrial, 600);
    }
    
    setTimeout(runNextTrial, 600);
}

// --- Vertex AI Model Monitoring Simulator ---
function initModelMonitorSimulator() {
    const slider = document.getElementById('drift-severity-slider');
    const label = document.getElementById('drift-val-label');
    const runBtn = document.getElementById('drift-simulate-btn');
    
    if (!slider || !runBtn) return;
    
    slider.addEventListener('input', () => {
        let val = parseFloat(slider.value);
        if (val === 0) label.textContent = "Baseline (0.00)";
        else if (val <= 0.3) label.textContent = `Minimal (${val.toFixed(2)})`;
        else if (val <= 0.6) label.textContent = `Moderate (${val.toFixed(2)})`;
        else label.textContent = `Critical (${val.toFixed(2)})`;
    });
    
    runBtn.addEventListener('click', () => {
        const severity = parseFloat(slider.value);
        runModelMonitoringSimulation(severity);
    });
}

function runModelMonitoringSimulation(severity) {
    const psiDisplay = document.getElementById('monitor-psi-display');
    const path = document.getElementById('monitor-line-path');
    const alertBox = document.getElementById('monitor-alert-box');
    const logConsole = document.getElementById('monitor-log-console');
    
    logConsole.innerHTML = `<span class="code-comment">// Querying prediction logs...</span>\n`;
    
    setTimeout(() => {
        let basePoints = [0.02, 0.03, 0.02, 0.04, 0.03, 0.05, 0.04];
        let lastPSI = 0.02 + severity * 0.45 + Math.random() * 0.05;
        let allPoints = [...basePoints, lastPSI];
        
        let pathD = "M 40 " + Math.round(150 - (allPoints[0]/0.5)*130);
        for (let i = 1; i < allPoints.length; i++) {
            let x = 40 + i * 60;
            let y = Math.round(150 - (allPoints[i]/0.5)*130);
            pathD += ` L ${x} ${y}`;
        }
        
        path.setAttribute('d', pathD);
        psiDisplay.textContent = lastPSI.toFixed(3);
        
        logConsole.innerHTML += `[SCAN] Checking 10,000 requests. Day 8 PSI = ${lastPSI.toFixed(3)}\n`;
        
        if (lastPSI >= 0.20) {
            path.classList.add('alert-active');
            alertBox.classList.add('active');
            psiDisplay.style.color = "var(--accent-pink)";
            logConsole.innerHTML += `<span style="color: var(--accent-pink);">[ALERT] Population Stability Index crossed 0.20. Skew detected.</span>\n`;
            logConsole.innerHTML += `[RETRAIN] Triggered Cloud Composer webhook to trigger retraining DAG.\n`;
        } else {
            path.classList.remove('alert-active');
            alertBox.classList.remove('active');
            psiDisplay.style.color = "var(--accent-blue)";
            logConsole.innerHTML += `<span style="color: var(--accent-green);">[HEALTHY] PSI is below the 0.20 drift alert threshold. Model operations stable.</span>\n`;
        }
        logConsole.scrollTop = logConsole.scrollHeight;
    }, 800);
}

// --- LLM Parameters Sandbox & Token Probability Visualizer ---
function initLLMSandbox() {
    const promptSelect = document.getElementById('llm-prompt-select');
    const tempSlider = document.getElementById('llm-temp-slider');
    const topKSlider = document.getElementById('llm-topk-slider');
    const topPSlider = document.getElementById('llm-topp-slider');
    
    const tempVal = document.getElementById('llm-temp-val');
    const topKVal = document.getElementById('llm-topk-val');
    const topPVal = document.getElementById('llm-topp-val');
    
    const generateBtn = document.getElementById('llm-generate-btn');
    
    if (!generateBtn) return;
    
    tempSlider.addEventListener('input', () => {
        tempVal.textContent = Number(tempSlider.value).toFixed(2);
        updateTokenProbabilityVisualization();
    });
    
    topKSlider.addEventListener('input', () => {
        topKVal.textContent = topKSlider.value;
        updateTokenProbabilityVisualization();
    });
    
    topPSlider.addEventListener('input', () => {
        topPVal.textContent = Number(topPSlider.value).toFixed(2);
        updateTokenProbabilityVisualization();
    });
    
    promptSelect.addEventListener('change', () => {
        updateTokenProbabilityVisualization();
    });
    
    updateTokenProbabilityVisualization();
    
    generateBtn.addEventListener('click', () => {
        const prompt = promptSelect.value;
        const temp = parseFloat(tempSlider.value);
        const topK = parseInt(topKSlider.value);
        const topP = parseFloat(topPSlider.value);
        runLLMGenerationSimulation(prompt, temp, topK, topP);
    });
}

function calculateProbabilities(promptText, temp, topK, topP) {
    const vocab = promptVocabulary[promptText] || [];
    if (vocab.length === 0) return [];
    
    let logits = vocab.map(v => {
        let baseLogit = Math.log(v.baseProb);
        let temperatureAdjustedLogit = temp <= 0.01 ? baseLogit * 100 : baseLogit / temp;
        return {
            token: v.token,
            logit: temperatureAdjustedLogit
        };
    });
    
    const exps = logits.map(l => Math.exp(l.logit));
    const sumExps = exps.reduce((a, b) => a + b, 0);
    let probs = logits.map((l, idx) => ({
        token: l.token,
        prob: exps[idx] / sumExps
    }));
    
    probs.sort((a, b) => b.prob - a.prob);
    
    probs = probs.map((item, idx) => {
        return {
            ...item,
            filteredByTopK: idx >= topK
        };
    });
    
    let cumulativeProb = 0;
    let cutoffReached = false;
    
    probs = probs.map((item, idx) => {
        if (item.filteredByTopK) {
            return { ...item, filteredByTopP: true };
        }
        
        if (cutoffReached) {
            return { ...item, filteredByTopP: true };
        }
        
        cumulativeProb += item.prob;
        if (cumulativeProb >= topP) {
            cutoffReached = true;
        }
        
        return { ...item, filteredByTopP: false };
    });
    
    const activeTokens = probs.filter(p => !p.filteredByTopK && !p.filteredByTopP);
    
    if (activeTokens.length > 0) {
        const activeSum = activeTokens.reduce((sum, t) => sum + t.prob, 0);
        probs = probs.map(p => {
            if (p.filteredByTopK || p.filteredByTopP) {
                return { ...p, displayProb: 0 };
            } else {
                return { ...p, displayProb: p.prob / activeSum };
            }
        });
    } else {
        probs = probs.map((p, idx) => ({
            ...p,
            displayProb: idx === 0 ? 1.0 : 0
        }));
    }
    
    return probs;
}

function updateTokenProbabilityVisualization() {
    const promptText = document.getElementById('llm-prompt-select').value;
    const temp = parseFloat(document.getElementById('llm-temp-slider').value);
    const topK = parseInt(document.getElementById('llm-topk-slider').value);
    const topP = parseFloat(document.getElementById('llm-topp-slider').value);
    
    const chartList = document.getElementById('llm-prob-chart-list');
    if (!chartList) return;
    
    const processedProbs = calculateProbabilities(promptText, temp, topK, topP);
    chartList.innerHTML = '';
    
    processedProbs.forEach(item => {
        const isFiltered = item.filteredByTopK || item.filteredByTopP;
        const displayPercent = Math.round(item.displayProb * 100);
        const originalPercent = Math.round(item.prob * 100);
        
        let filterLabel = '';
        if (item.filteredByTopK) filterLabel = ' [Filtered by Top-K]';
        else if (item.filteredByTopP) filterLabel = ' [Filtered by Top-P]';
        
        const barItem = document.createElement('div');
        barItem.className = `prob-bar-item`;
        
        const opacityStyle = isFiltered ? 'opacity: 0.35;' : '';
        
        barItem.innerHTML = `
            <div class="prob-bar-label" style="${opacityStyle}">
                <span>"${item.token}"${filterLabel}</span>
                <span>${isFiltered ? '0%' : `${displayPercent}%`} <span style="font-size: 0.65rem; opacity: 0.6;">(Raw: ${originalPercent}%)</span></span>
            </div>
            <div class="prob-bar-track" style="${opacityStyle}">
                <div class="prob-bar-fill ${isFiltered ? '' : 'selected'}" style="width: ${isFiltered ? '0%' : `${displayPercent}%`}"></div>
            </div>
        `;
        chartList.appendChild(barItem);
    });
}

function runLLMGenerationSimulation(prompt, temp, topK, topP) {
    const chatHistory = document.getElementById('sandbox-chat-history');
    
    const userMsg = document.createElement('div');
    userMsg.className = 'chat-bubble user';
    userMsg.textContent = prompt;
    chatHistory.appendChild(userMsg);
    
    const processedProbs = calculateProbabilities(prompt, temp, topK, topP);
    const activeTokens = processedProbs.filter(p => p.displayProb > 0);
    
    let nextToken = "unknown";
    if (activeTokens.length > 0) {
        const rand = Math.random();
        let runningSum = 0;
        for (let t of activeTokens) {
            runningSum += t.displayProb;
            if (rand <= runningSum) {
                nextToken = t.token;
                break;
            }
        }
    }
    
    const assistantMsg = document.createElement('div');
    assistantMsg.className = 'chat-bubble assistant';
    assistantMsg.innerHTML = `<em>Model is calculating next token...</em>`;
    chatHistory.appendChild(assistantMsg);
    chatHistory.scrollTop = chatHistory.scrollHeight;
    
    setTimeout(() => {
        let fullOutputText = `"${nextToken}"`;
        
        if (prompt === "Machine learning is") {
            if (nextToken === "powerful") fullOutputText = `<strong>${nextToken}</strong> when paired with reliable operations pipelines like MLOps.`;
            else if (nextToken === "fun") fullOutputText = `<strong>${nextToken}</strong> to prototype, but running it in production requires strict controls!`;
            else if (nextToken === "difficult") fullOutputText = `<strong>${nextToken}</strong> to operationalize at scale without a managed platform like Vertex AI.`;
            else fullOutputText = `<strong>${nextToken}</strong> and continues to reshape software engineering models.`;
        } else if (prompt === "Generative AI will") {
            if (nextToken === "transform") fullOutputText = `<strong>${nextToken}</strong> how software products are built, shifting focus towards cognitive user interfaces.`;
            else if (nextToken === "automate") fullOutputText = `<strong>${nextToken}</strong> routine text summaries and document analyses inside enterprise systems.`;
            else if (nextToken === "replace") fullOutputText = `<strong>${nextToken}</strong> manual template creation for routine coding tasks and documents.`;
            else fullOutputText = `<strong>${nextToken}</strong> the business workflows of digital operations.`;
        } else if (prompt === "Vertex AI Feature Store solves") {
            if (nextToken === "skew") fullOutputText = `train-serve <strong>${nextToken}</strong> by unifying features across offline training exports and online inference calls.`;
            else if (nextToken === "latency") fullOutputText = `real-time inference <strong>${nextToken}</strong> bottlenecks by serving features from low-latency databases (<= 15ms).`;
            else if (nextToken === "leakage") fullOutputText = `historical data <strong>${nextToken}</strong> problems using built-in point-in-time snapshots during training.`;
            else fullOutputText = `<strong>${nextToken}</strong> and duplicates across engineering teams.`;
        } else {
            fullOutputText = `<strong>${nextToken}</strong> token sequence elements efficiently across high-dimension coordinate nodes.`;
        }
        
        assistantMsg.innerHTML = `
            <div>${fullOutputText}</div>
            <div style="font-size: 0.7rem; color: var(--text-muted); margin-top: 8px;">
                Parameters used: Temp = ${temp.toFixed(2)} | Top-K = ${topK} | Top-P = ${topP.toFixed(2)}
            </div>
        `;
        chatHistory.scrollTop = chatHistory.scrollHeight;
    }, 1200);
}

// --- Practice Exam Logic (25 Mapped Questions) ---
function initQuiz() {
    const startBtn = document.getElementById('quiz-start-btn');
    if (!startBtn) return;
    
    startBtn.addEventListener('click', () => {
        document.getElementById('quiz-landing-container').style.display = 'none';
        document.getElementById('quiz-active-container').style.display = 'block';
        
        state.quiz.currentQuestion = 0;
        state.quiz.score = 0;
        state.quiz.userAnswers = [];
        state.quiz.answered = false;
        
        renderQuizQuestion();
        startQuizTimer();
    });
}

function startQuizTimer() {
    if (state.quiz.timerInterval) clearInterval(state.quiz.timerInterval);
    
    state.quiz.timer = 0;
    const timerText = document.getElementById('quiz-timer-text');
    
    state.quiz.timerInterval = setInterval(() => {
        state.quiz.timer++;
        
        let minutes = Math.floor(state.quiz.timer / 60);
        let seconds = state.quiz.timer % 60;
        
        minutes = minutes < 10 ? '0' + minutes : minutes;
        seconds = seconds < 10 ? '0' + seconds : seconds;
        
        if (timerText) {
            timerText.textContent = `${minutes}:${seconds}`;
        }
    }, 1000);
}

function renderQuizQuestion() {
    const qIdx = state.quiz.currentQuestion;
    const q = quizQuestions[qIdx];
    
    state.quiz.selectedChoice = null;
    state.quiz.answered = false;
    
    const questionNumEl = document.getElementById('quiz-question-number');
    if (questionNumEl) {
        questionNumEl.textContent = `Question ${qIdx + 1} of ${quizQuestions.length}`;
    }
    
    const qTextEl = document.getElementById('quiz-question-text');
    if (qTextEl) {
        qTextEl.textContent = q.question;
    }
    
    const choicesContainer = document.getElementById('quiz-choices-container');
    choicesContainer.innerHTML = '';
    
    const labelMapping = ['A', 'B', 'C', 'D'];
    
    q.choices.forEach((choice, idx) => {
        const choiceItem = document.createElement('div');
        choiceItem.className = 'quiz-choice-item';
        choiceItem.innerHTML = `
            <div class="choice-marker">${labelMapping[idx]}</div>
            <span>${choice}</span>
        `;
        
        choiceItem.addEventListener('click', () => {
            if (state.quiz.answered) return;
            
            document.querySelectorAll('.quiz-choice-item').forEach(c => {
                c.classList.remove('selected');
            });
            
            choiceItem.classList.add('selected');
            state.quiz.selectedChoice = idx;
            document.getElementById('quiz-submit-btn').disabled = false;
        });
        
        choicesContainer.appendChild(choiceItem);
    });
    
    const submitBtn = document.getElementById('quiz-submit-btn');
    submitBtn.disabled = true;
    submitBtn.style.display = 'block';
    
    const nextBtn = document.getElementById('quiz-next-btn');
    nextBtn.style.display = 'none';
    
    const explanationEl = document.getElementById('quiz-explanation-box');
    explanationEl.className = 'quiz-explanation';
    explanationEl.innerHTML = '';
}

function checkAnswer() {
    const qIdx = state.quiz.currentQuestion;
    const q = quizQuestions[qIdx];
    const selected = state.quiz.selectedChoice;
    
    if (selected === null || state.quiz.answered) return;
    
    state.quiz.answered = true;
    
    const isCorrect = selected === q.correctAnswer;
    if (isCorrect) {
        state.quiz.score++;
    }
    
    state.quiz.userAnswers.push({
        questionIndex: qIdx,
        selectedChoice: selected,
        isCorrect: isCorrect
    });
    
    const choiceItems = document.querySelectorAll('.quiz-choice-item');
    choiceItems.forEach((item, idx) => {
        item.classList.add('disabled');
        if (idx === q.correctAnswer) {
            item.classList.add('correct');
        } else if (idx === selected) {
            item.classList.add('incorrect');
        }
    });
    
    const explanationEl = document.getElementById('quiz-explanation-box');
    explanationEl.innerHTML = q.explanation;
    if (isCorrect) {
        explanationEl.classList.add('active-correct');
    } else {
        explanationEl.classList.add('active-incorrect');
    }
    
    document.getElementById('quiz-submit-btn').style.display = 'none';
    const nextBtn = document.getElementById('quiz-next-btn');
    nextBtn.style.display = 'block';
    
    if (qIdx === quizQuestions.length - 1) {
        nextBtn.textContent = "Finish & View Results";
        nextBtn.onclick = () => {
            finishQuiz();
        };
    } else {
        nextBtn.textContent = "Next Question";
        nextBtn.onclick = () => {
            state.quiz.currentQuestion++;
            renderQuizQuestion();
        };
    }
    
    updateGlobalProgress();
}

function finishQuiz() {
    if (state.quiz.timerInterval) clearInterval(state.quiz.timerInterval);
    
    document.getElementById('quiz-active-container').style.display = 'none';
    const resultsContainer = document.getElementById('quiz-results-container');
    resultsContainer.style.display = 'block';
    
    const totalQuestions = quizQuestions.length;
    const finalScore = state.quiz.score;
    const percent = Math.round((finalScore / totalQuestions) * 100);
    
    document.getElementById('quiz-score-val').textContent = finalScore;
    document.getElementById('quiz-total-val').textContent = `out of ${totalQuestions} questions`;
    
    const radialCircle = document.getElementById('quiz-radial-progress');
    if (radialCircle) {
        radialCircle.style.setProperty('--percent', `${percent}%`);
    }
    
    const commentaryEl = document.getElementById('quiz-commentary');
    if (percent >= 70) {
        commentaryEl.innerHTML = `<span style="color: var(--accent-green); font-weight:700;">PASS! Score: ${percent}%</span><br>Congratulations! You have demonstrated readiness across all syllabus areas. You are well prepared for the GCP Professional Machine Learning Engineer certification exam.`;
    } else {
        commentaryEl.innerHTML = `<span style="color: var(--accent-pink); font-weight:700;">FAIL! Score: ${percent}%</span><br>The passing score threshold is 70%. We recommend studying the course slides for custom containers, pipeline orchestration, and model monitoring alerts, and retaking the practice test.`;
    }
    
    const reviewList = document.getElementById('quiz-review-list');
    if (reviewList) {
        reviewList.innerHTML = '';
        state.quiz.userAnswers.forEach((ans, idx) => {
            const q = quizQuestions[ans.questionIndex];
            const labels = ['A', 'B', 'C', 'D'];
            
            const reviewItem = document.createElement('div');
            reviewItem.style.background = 'rgba(255,255,255,0.02)';
            reviewItem.style.border = '1px solid var(--border-color)';
            reviewItem.style.padding = '16px';
            reviewItem.style.borderRadius = '8px';
            reviewItem.style.marginBottom = '12px';
            reviewItem.style.textAlign = 'left';
            
            reviewItem.innerHTML = `
                <div style="font-weight: 600; display:flex; justify-content:space-between; margin-bottom:8px;">
                    <span>Question ${idx + 1}</span>
                    <span style="color: ${ans.isCorrect ? 'var(--accent-green)' : 'var(--accent-pink)'}">
                        ${ans.isCorrect ? 'Correct ✓' : 'Incorrect ✗'}
                    </span>
                </div>
                <div style="font-size: 0.9rem; color: var(--text-secondary); margin-bottom: 8px;">${q.question}</div>
                <div style="font-size: 0.85rem; color: var(--text-muted);">
                    Your selection: <strong>${labels[ans.selectedChoice]}</strong> | Correct: <strong style="color: var(--accent-green);">${labels[q.correctAnswer]}</strong>
                </div>
            `;
            reviewList.appendChild(reviewItem);
        });
    }
    
    const retakeBtn = document.getElementById('quiz-retake-btn');
    if (retakeBtn) {
        retakeBtn.onclick = () => {
            resultsContainer.style.display = 'none';
            document.getElementById('quiz-active-container').style.display = 'block';
            state.quiz.currentQuestion = 0;
            state.quiz.score = 0;
            state.quiz.userAnswers = [];
            state.quiz.answered = false;
            renderQuizQuestion();
            startQuizTimer();
        };
    }
}
