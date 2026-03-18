#!/bin/bash
# Deploy MS. READ Content Engine to Google Cloud Run
# Usage: ./deploy.sh
#
# Prerequisites:
#   1. Install gcloud CLI: https://cloud.google.com/sdk/docs/install
#   2. Run: gcloud auth login
#   3. Run: gcloud config set project YOUR_PROJECT_ID

set -e

PROJECT_ID=$(gcloud config get-value project 2>/dev/null)
REGION="asia-southeast1"  # Singapore — closest to Malaysia
SERVICE_NAME="msread-content-engine"
GOOGLE_AI_API_KEY="${GOOGLE_AI_API_KEY:-AIzaSyCAK7AsNYktX-eggwtOcLG-nwH-jPAqako}"

if [ -z "$PROJECT_ID" ] || [ "$PROJECT_ID" = "(unset)" ]; then
    echo "Error: No GCP project set."
    echo "Run: gcloud config set project YOUR_PROJECT_ID"
    exit 1
fi

echo ""
echo "  Deploying MS. READ Content Engine to Cloud Run"
echo "  Project:  $PROJECT_ID"
echo "  Region:   $REGION"
echo ""

# Build and deploy in one step
gcloud run deploy "$SERVICE_NAME" \
    --source . \
    --region "$REGION" \
    --platform managed \
    --allow-unauthenticated \
    --set-env-vars "GOOGLE_AI_API_KEY=$GOOGLE_AI_API_KEY" \
    --memory 1Gi \
    --timeout 600 \
    --min-instances 0 \
    --max-instances 3

# Get the URL
URL=$(gcloud run services describe "$SERVICE_NAME" --region "$REGION" --format 'value(status.url)')

echo ""
echo "  ================================================"
echo "  Deployed!"
echo "  URL: $URL"
echo "  ================================================"
echo ""
echo "  Share this URL with your colleagues."
echo ""
