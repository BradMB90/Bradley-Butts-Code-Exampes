/*
Script Name: SB_EnemyController.cs
Author: Bradley M. Butts
Last Modified: 10-19-2015
Description: This script handles all of the enemy's AI components. All of the major components must be entered inside of the inspector
             in the Unity development window. The AI mechanics include: pathfinding, automated movement based on the path found,
             checking if the player is a certain number of tiles away from the enemy based on the direction the enemy is looking at.
*/

using UnityEngine;
using System.Collections;
using Pathfinding;
using Rotorz.Tile;
using Rotorz.Tile.Internal;

public class SB_EnemyController : MonoBehaviour
{

    private GameObject player;
    public SB_GameController gameCon;
    public TileSystem tileSystem;
    public TileIndex startTile, endTile;
    public int lineOfSight;
    private Vector3 startPos, endPos;
    public Seeker seeker;
    public CharacterController controller;
    private Path path;
    public float speed = 1;
    private Vector3[] pathArray;
    private int currentPos;
    public bool up, down, right, left, moved, faceUp, faceDown, faceRight, faceLeft, isAlarmed;
    public bool to, from, updatedPos;

    // Use this for initialization
    void Start()
    {
        GetComponent<Renderer>().shadowCastingMode = UnityEngine.Rendering.ShadowCastingMode.On;
        GetComponent<Renderer>().receiveShadows = true;
        player = GameObject.Find("SB_Player");
        tileSystem = player.GetComponent<SB_PlayerController>().tileSystem;
        startPos = tileSystem.GetTile(startTile).gameObject.transform.position;
        startPos.y = .74f;
        transform.position = startPos;
        endPos = tileSystem.GetTile(endTile).gameObject.transform.position;
        endPos.y = .74f;
        gameCon = player.GetComponent<SB_GameController>();
        seeker = GetComponent<Seeker>();
        currentPos = 0;
        controller = GetComponent<CharacterController>();
        //A* function that generates the closest path from startPos to endPos.
        seeker.StartPath(startPos, endPos, OnPathComplete);
        to = true; from = false; updatedPos = false; moved = true; isAlarmed = false;
    }

    // Update is called once per frame. Looks for the player whenever the player's turn is done
    // Only moves if the player's count is the same as the game's count (which means the player has successfully moved)
    void Update()
    {
        if (!gameCon.isLevelPaused)
        {
            if (!isAlarmed)
            {
                if (!player.GetComponent<SB_PlayerController>().moving)
                {
                    LookForPlayer();
                }
                if (gameCon.playerCount > gameCon.enemyCount)
                {
                    moved = false;
                    MoveToNextLoc();
                }
                if (gameCon.playerCount == gameCon.enemyCount)
                {
                    moved = true;
                }
            }
            else
            {
                LookForPlayer();
            }
        }
    }

    //Function gets the direction the enemy should be facing based on the next tile
    void GetDirection(Vector3 start, Vector3 end)
    {
        if ((start.x > end.x) && (Mathf.Approximately(start.z, end.z)) && !V3Equal(start, end))
        {
            up = false; down = false; right = false; left = true;
            faceUp = false; faceDown = false; faceRight = false; faceLeft = true;
        }
        else if ((start.x < end.x) && (Mathf.Approximately(start.z, end.z)) && !V3Equal(start, end))
        {
            up = false; down = false; right = true; left = false;
            faceUp = false; faceDown = false; faceRight = true; faceLeft = false;
        }
        else if ((start.z > end.z) && (Mathf.Approximately(start.x, end.x)) && !V3Equal(start, end))
        {
            up = false; down = true; right = false; left = false;
            faceUp = false; faceDown = true; faceRight = false; faceLeft = false;
        }
        else if ((start.z < end.z) && (Mathf.Approximately(start.x, end.x)) && !V3Equal(start, end))
        {
            up = true; down = false; right = false; left = false;
            faceUp = true; faceDown = false; faceRight = false; faceLeft = false;
        }
    }

    //Function changes the Y value of the path returned to a specified value. This is to keep the enemy's Y value consistant since the path generated has all Y values
    //set to 0 by default; which would have the player walking inside the ground
    public void ChangeYValue(Vector3[] path)
    {
        for (int i = 0; i < path.Length; i++)
        {
            Vector3 temp = path[i];
            temp.y = .74f;
            path[i] = temp;
        }
    }

    //Once the A* path is complete; it takes the path generated and stores it into a Vector3 array for the enemy to follow
    void OnPathComplete(Path p)
    {
        if (!p.error)
        {
            path = p;
            pathArray = path.vectorPath.ToArray();
            ChangeYValue(pathArray);
            GetDirection(startPos, pathArray[currentPos + 1]); moved = false; moved = true;
        }
    }

    //Based on the direction the enemy is facing; the enemy looks for the player a set number of tiles above, below, left, or right. If an object is in the line of sight of the enemy
    //The LookForPlayer function ends to prevent the enemy from seeing the player through objects like trees or rocks.
    public void LookForPlayer()
    {
        if (faceUp)
        {
            try
            {
                for (int i = 1; i <= lineOfSight; i++)
                {
                    TileIndex tempTile = tileSystem.ClosestTileIndexFromWorld(transform.position);
                    if (tempTile.row - i >= 0)
                    {
                        SB_TileCheck tempTileCheck = tileSystem.GetTile((tempTile.row - i), tempTile.column).gameObject.GetComponent<SB_TileCheck>();
                        if (tempTileCheck.occupied == true)
                        {
                            if (tempTileCheck.occupier.tag == "Player")
                            {
                                gameCon.isLevelLost = true;
                            }
                            else
                            {
                                i = lineOfSight + 1;
                            }
                        }
                    }
                }
            }
            catch { }
        }
        if (faceDown)
        {
            try
            {
                for (int i = 1; i <= lineOfSight; i++)
                {
                    TileIndex tempTile = tileSystem.ClosestTileIndexFromWorld(transform.position);
                    if (tempTile.row + i <= tileSystem.RowCount - 1)
                    {
                        SB_TileCheck tempTileCheck = tileSystem.GetTile((tempTile.row + i), tempTile.column).gameObject.GetComponent<SB_TileCheck>();
                        if (tempTileCheck.occupied == true)
                        {
                            if (tempTileCheck.occupier.tag == "Player")
                            {
                                gameCon.isLevelLost = true;
                            }
                            else
                            {
                                i = lineOfSight + 1;
                            }
                        }
                    }
                }
            }
            catch { }
        }
        if (faceRight)
        {
            try
            {
                for (int i = 1; i <= lineOfSight; i++)
                {
                    TileIndex tempTile = tileSystem.ClosestTileIndexFromWorld(transform.position);
                    if (tempTile.column + i <= tileSystem.ColumnCount - 1)
                    {
                        SB_TileCheck tempTileCheck = tileSystem.GetTile(tempTile.row, (tempTile.column + i)).gameObject.GetComponent<SB_TileCheck>();
                        if (tempTileCheck.occupied == true)
                        {
                            if (tempTileCheck.occupier.tag == "Player")
                            {
                                gameCon.isLevelLost = true;
                            }
                            else
                            {
                                i = lineOfSight + 1;
                            }
                        }
                    }
                }
            }
            catch { }
        }
        if (faceLeft)
        {
            try
            {
                for (int i = 1; i <= lineOfSight; i++)
                {
                    TileIndex tempTile = tileSystem.ClosestTileIndexFromWorld(transform.position);
                    if (tempTile.column - i <= 0)
                    {
                        SB_TileCheck tempTileCheck = tileSystem.GetTile(tempTile.row, (tempTile.column - i)).gameObject.GetComponent<SB_TileCheck>();
                        if (tempTileCheck.occupied == true)
                        {
                            if (tempTileCheck.occupier.tag == "Player")
                            {
                                gameCon.isLevelLost = true;
                            }
                            else
                            {
                                i = lineOfSight + 1;
                            }
                        }
                    }
                }
            }
            catch { }
        }
    }

    //Script that we use to check if the enemy's current location is close enough to the next tiles location to end the movement script
    public bool V3Equal(Vector3 a, Vector3 b)
    {
        return Vector3.SqrMagnitude(a - b) < 0.000001;
    }

    //Function moves the enemy to the next tile in the path array we generated earlier. If the enemy reaches the end of the array (or the beginning of the array if the enemy is returning
    //the function also changes the to and from value respectfully. 
    public void MoveToNextLoc()
    {
        if (!moved)
        {
            if (to)
            {
                if ((currentPos != pathArray.Length - 1) && (tileSystem.GetTile(tileSystem.ClosestTileIndexFromWorld(pathArray[currentPos + 1])).gameObject.GetComponent<SB_TileCheck>().occupier.tag == "Enemy") && (tileSystem.GetTile(tileSystem.ClosestTileIndexFromWorld(pathArray[currentPos + 1])).gameObject.GetComponent<SB_TileCheck>().occupier != gameObject))
                {
                    moved = true;
                    updatedPos = false;
                    gameCon.enemyDone++;
                }
                else
                {
                    if (!updatedPos && currentPos != pathArray.Length - 1)
                    {
                        currentPos++;
                        updatedPos = true;
                    }
                    if ((!V3Equal(pathArray[currentPos], transform.position)) && (currentPos <= pathArray.Length - 1))
                    {
                        Vector3 dir = (pathArray[currentPos] - transform.position).normalized;
                        dir *= Time.fixedDeltaTime * speed;
                        controller.Move(dir);
                        GetDirection(transform.position, pathArray[currentPos]);
                    }
                    if (V3Equal(pathArray[currentPos], transform.position))
                    {
                        updatedPos = false;
                        moved = true;
                        gameCon.enemyDone++;
                    }
                    if (currentPos == pathArray.Length - 1)
                    {
                        moved = true;
                        to = false;
                        from = true;
                    }
                }
            }
            if (from)
            {
                if ((currentPos != 0) && (tileSystem.GetTile(tileSystem.ClosestTileIndexFromWorld(pathArray[currentPos - 1])).gameObject.GetComponent<SB_TileCheck>().occupier.tag == "Enemy") && (tileSystem.GetTile(tileSystem.ClosestTileIndexFromWorld(pathArray[currentPos - 1])).gameObject.GetComponent<SB_TileCheck>().occupier != gameObject))
                {
                    moved = true;
                    updatedPos = false;
                    gameCon.enemyDone++;
                }
                else
                {
                    if (!updatedPos && currentPos != 0)
                    {
                        currentPos--;
                        updatedPos = true;
                    }
                    if ((!V3Equal(pathArray[currentPos], transform.position)) && (currentPos >= 0))
                    {
                        Vector3 dir = (pathArray[currentPos] - transform.position).normalized;
                        dir *= Time.fixedDeltaTime * speed;
                        controller.Move(dir);
                        GetDirection(transform.position, pathArray[currentPos]);
                    }
                    if (V3Equal(pathArray[currentPos], transform.position))
                    {
                        updatedPos = false;
                        moved = true;
                        gameCon.enemyDone++;
                    }
                    if (currentPos == 0)
                    {
                        moved = true;
                        to = true;
                        from = false;
                    }
                }
            }
        }
    }
}
