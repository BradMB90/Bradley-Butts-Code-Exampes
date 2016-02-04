/*
Script Name: SB_PlayerController.cs
Author: Bradley M. Butts
Last Modified: 10-19-2015
Description: This script handles all the player's functions inside of the game.
             These functions include: Initial setup, showing which direction the player is moving to handle animations,
             updating the UI buttons to indicate where the player can or cannot move, checking which tiles
             are and are not valid to move to, and movement commands to move to the next selected tile.
*/

using UnityEngine;
using UnityEngine.UI;
using UnityEngine.EventSystems;
using System.Collections;
using System.Collections.Generic;
using Rotorz.Tile;
using Rotorz.Tile.Internal;

public class SB_PlayerController : MonoBehaviour {

	public TileSystem tileSystem;
    //Current tile the player is on and the next tile that the player selects to move to
	public TileIndex currentTile, nextTile;
    //XYZ data for the selected tiles above
	private Vector3 currentLoc, nextLoc;
	private CharacterController controller;
	private float speed = 1;
	public bool moving, up, down, left, right;
	public Button upButton, downButton, rightButton, leftButton;
	private TileData upTile, downTile, rightTile, leftTile;
	private SB_GameController gameCon;
    //Icon that shows whenever there is an area the player can interact with
    public GameObject interactiveIcon;

	// Use this for initialization
	void Start () {
		gameObject.tag = "Player";
        GetComponent<Renderer>().shadowCastingMode = UnityEngine.Rendering.ShadowCastingMode.On;
        GetComponent<Renderer> ().receiveShadows = true;
		controller = gameObject.GetComponent<CharacterController> ();
		gameCon = gameObject.GetComponent<SB_GameController> ();
		moving = false;
		currentTile = gameCon.startTile;
		currentLoc = tileSystem.GetTile (currentTile).gameObject.transform.position;
		currentLoc.y = .75f;
		transform.position = currentLoc;
	}
	
	// Update is called once per frame
	void Update () {
		UpdateButtons ();
        //While the game is not paused and the player is not moving, the script will update the UI to show which directions the
        //player can move, it'll also turn off all direction movements to render the "idle" animation
		if (!gameCon.isLevelPaused) {
			if(!moving){
				GetValidTiles ();
				up = false;
				down = false;
				right = false;
				left = false;
			}
            //While moving is true, script will run the function MoveToNextTile until the player
            //is on the selected tile
			if (moving) {
				MoveToNextTile();
				//ResetTiles ();
			}
            //If the player's current tile is equal to the end tile; the player has won the level
			if(currentTile == gameCon.endTile){
				gameCon.isLevelWon = true;
			}
		}
	}

    //Script below is current obsolete
	/*void ResetTiles(){
		for (int i = 0; i < tileSystem.RowCount; i++) {
			for(int j = 0; j < tileSystem.ColumnCount; j++){
				tileSystem.GetTile (i,j).gameObject.GetComponent<SB_TileCheck>().valid = false;
			}
		}
	}*/


    //Functions handles the UI for the player's input. These UI components include the Interact button and movement buttons.
	void UpdateButtons(){
        if (gameCon.canInteract)
        {
            interactiveIcon.SetActive(true);
        }
        else
        {
            interactiveIcon.SetActive(false);
        }
		if (!gameCon.isLevelPaused) {
            //While moving, we turn off interaction with all movement buttons. This will make the player wait for their character to stop moving as well as the enemies to stop moving.
			if (moving) {
				upButton.interactable = false;
				downButton.interactable = false;
				rightButton.interactable = false;
				leftButton.interactable = false;
			}
            //During the player's turn, the below code turns on and off the direction movement buttons based on whether or not the tile is occupied and/or valid
			if (!moving && (gameCon.playerCount == gameCon.enemyCount)) {
				try{
					if(upTile.gameObject.GetComponent<SB_TileCheck>().valid && currentTile.row != 0){
						upButton.interactable = true;
					}
					if(!upTile.gameObject.GetComponent<SB_TileCheck>().valid || currentTile.row == 0 || upTile.gameObject.GetComponent<SB_TileCheck>().occupied){
						upButton.interactable = false;
					}
					if(downTile.gameObject.GetComponent<SB_TileCheck>().valid && currentTile.row != tileSystem.RowCount-1){
						downButton.interactable = true;
					}
					if(!downTile.gameObject.GetComponent<SB_TileCheck>().valid || currentTile.row == tileSystem.RowCount-1 || downTile.gameObject.GetComponent<SB_TileCheck>().occupied){
						downButton.interactable = false;
					}
					if(rightTile.gameObject.GetComponent<SB_TileCheck>().valid && currentTile.column != tileSystem.ColumnCount-1){
						rightButton.interactable = true;
					}
					if(!rightTile.gameObject.GetComponent<SB_TileCheck>().valid || currentTile.column == tileSystem.ColumnCount-1 || rightTile.gameObject.GetComponent<SB_TileCheck>().occupied){
						rightButton.interactable = false;
					}
					if(leftTile.gameObject.GetComponent<SB_TileCheck>().valid && currentTile.column != 0){
						leftButton.interactable = true;
					}
					if(!leftTile.gameObject.GetComponent<SB_TileCheck>().valid || currentTile.column == 0 || leftTile.gameObject.GetComponent<SB_TileCheck>().occupied){
						leftButton.interactable = false;
					}
				}
				catch{}
			}
		}
		else{
			upButton.interactable = false;
			downButton.interactable = false;
			rightButton.interactable = false;
			leftButton.interactable = false;
		}
	}

    //Public script attached to the Up button to move the player to the tile that is located one row above the current one
	public void GoUp(){
		if (upButton.interactable) {
			if (upTile.gameObject.GetComponent<SB_TileCheck> ().valid && gameCon.playerCount < gameCon.gameCount) {
				GetNextTile (tileSystem.ClosestTileIndexFromWorld (upTile.gameObject.transform.position));
				up = true;
			}
		}
	}

    //Public script attached to the Down button to move the player to the tile that is located one row below the current one
    public void GoDown(){
		if (downButton.interactable) {
			if (downTile.gameObject.GetComponent<SB_TileCheck> ().valid && gameCon.playerCount < gameCon.gameCount) {
				GetNextTile (tileSystem.ClosestTileIndexFromWorld (downTile.gameObject.transform.position));
				down = true;
			}
		}
	}

    //Public script attached to the Right button to move the player to the tile that is located one column right to the current one
    public void GoRight(){
		if (rightButton.interactable) {
			if (rightTile.gameObject.GetComponent<SB_TileCheck> ().valid && gameCon.playerCount < gameCon.gameCount) {
				GetNextTile (tileSystem.ClosestTileIndexFromWorld (rightTile.gameObject.transform.position));
				right = true;
			}
		}
	}

    //Public script attached to the Left button to move the player to the tile that is located one column left to the current one
    public void GoLeft(){
		if (leftButton.interactable) {
			if (leftTile.gameObject.GetComponent<SB_TileCheck> ().valid && gameCon.playerCount < gameCon.gameCount) {
				GetNextTile (tileSystem.ClosestTileIndexFromWorld (leftTile.gameObject.transform.position));
				left = true;
			}
		}
	}

    //Makes the user's select tile the next tile to move to, and changes the status of moving to true
	void GetNextTile(TileIndex next){
		nextTile = next;
		nextLoc = tileSystem.GetTile (nextTile).gameObject.transform.position;
		moving = true;
	}

    //Script that checks to see if the tiles surrounding the player (Up, Down, Left, Right) are valid tiles to move to
	void GetValidTiles(){
		try{
			upTile = tileSystem.GetTile (currentTile.row - 1, currentTile.column);
		}
		catch{}
		try{
			downTile = tileSystem.GetTile (currentTile.row + 1, currentTile.column);
		}
		catch{}
		try{
			rightTile = tileSystem.GetTile (currentTile.row, currentTile.column + 1);
		}
		catch{}
		try{
			leftTile = tileSystem.GetTile (currentTile.row, currentTile.column - 1);
		}
		catch{}

		try{
			if (upTile.gameObject.GetComponent<SB_TileCheck> ().occupied && upTile.gameObject.GetComponent<SB_TileCheck> ().occupier.tag != "Player") {
				upTile.gameObject.GetComponent<SB_TileCheck>().valid = false;
			}
			else{
				upTile.gameObject.GetComponent<SB_TileCheck>().valid = true;
			}
			if (downTile.gameObject.GetComponent<SB_TileCheck> ().occupied && downTile.gameObject.GetComponent<SB_TileCheck> ().occupier.tag != "Player") {
				downTile.gameObject.GetComponent<SB_TileCheck>().valid = false;
			}
			else{
				downTile.gameObject.GetComponent<SB_TileCheck>().valid = true;
			}
			if (rightTile.gameObject.GetComponent<SB_TileCheck> ().occupied && rightTile.gameObject.GetComponent<SB_TileCheck> ().occupier.tag != "Player") {
				rightTile.gameObject.GetComponent<SB_TileCheck>().valid = false;
			}
			else{
				rightTile.gameObject.GetComponent<SB_TileCheck>().valid = true;
			}
			if (leftTile.gameObject.GetComponent<SB_TileCheck> ().occupied && leftTile.gameObject.GetComponent<SB_TileCheck> ().occupier.tag != "Player") {
				leftTile.gameObject.GetComponent<SB_TileCheck>().valid = false;
			}
			else{
				leftTile.gameObject.GetComponent<SB_TileCheck>().valid = true;
			}
		}
		catch{}
	}

    //Script that we use to check if the player's current location is close enough to the next tiles location to end the movement script
	public bool V3Equal(Vector3 a, Vector3 b){
		return Vector3.SqrMagnitude(a-b) < 0.001;
	}

    //Script that moves the player to the next selected tile. 
	void MoveToNextTile(){
		if (V3Equal(transform.position, nextLoc)) {
			transform.position = nextLoc;
			currentTile = nextTile;
			moving = false;
			gameCon.playerCount++;
		}
		nextLoc.y = .75f;
		Vector3 moveDiff = nextLoc - transform.position;
		Vector3 moveDir = moveDiff.normalized * speed * Time.deltaTime;
		if(moveDir.sqrMagnitude < moveDiff.sqrMagnitude){
			controller.Move (moveDir);
		}
		else{
			controller.Move (moveDiff);
		}
	}
}
